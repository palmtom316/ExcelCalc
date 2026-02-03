using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;

namespace ExcelSmartCostAddIn
{
    internal static class QuantityRecalc
    {
        private const string COL_EXPR = "Calc_Expr";
        private const string COL_QTY = "Quantity";
        private const string COL_ROWTYPE = "RowType";
        private const string COL_NAME = "Name";
        private const string COL_UNIT = "Unit";

        public static async Task RecalculateActiveSheetAsync()
        {
            var app = Globals.ThisAddIn.Application;
            Excel.Worksheet ws = app.ActiveSheet;
            var headers = ReadHeaders(ws);
            if (!headers.ContainsKey(COL_EXPR) || !headers.ContainsKey(COL_QTY))
                throw new Exception("Sheet missing required headers: Calc_Expr, Quantity");

            int exprCol = headers[COL_EXPR];
            int qtyCol = headers[COL_QTY];
            int rowTypeCol = headers.ContainsKey(COL_ROWTYPE) ? headers[COL_ROWTYPE] : -1;
            int nameCol = headers.ContainsKey(COL_NAME) ? headers[COL_NAME] : -1;
            int unitCol = headers.ContainsKey(COL_UNIT) ? headers[COL_UNIT] : -1;

            bool prevSU = app.ScreenUpdating;
            Excel.XlCalculation prevCalc = app.Calculation;
            app.ScreenUpdating = false;
            app.Calculation = Excel.XlCalculation.xlCalculationManual;

            try
            {
                int lastRow = ws.UsedRange.Rows.Count;
                if (lastRow < 2) return;

                if (rowTypeCol > 0)
                {
                    RemoveCompOutputRows(ws, rowTypeCol, lastRow);
                    lastRow = ws.UsedRange.Rows.Count;
                    if (lastRow < 2) return;
                }

                Excel.Range exprRange = ws.Range[ws.Cells[2, exprCol], ws.Cells[lastRow, exprCol]];
                object[,] exprVals = Ensure2D(exprRange.Value2, lastRow - 1, 1);

                object[,] unitVals = null;
                object[,] unitOut = null;
                if (unitCol > 0)
                {
                    Excel.Range unitRange = ws.Range[ws.Cells[2, unitCol], ws.Cells[lastRow, unitCol]];
                    unitVals = Ensure2D(unitRange.Value2, lastRow - 1, 1);
                    unitOut = new object[lastRow - 1, 1];
                    for (int i = 1; i <= lastRow - 1; i++)
                    {
                        unitOut[i - 1, 0] = unitVals[i, 1];
                    }
                }

                using (var client = EngineClient.Create())
                {
                    var qtyOut = new object[lastRow - 1, 1];
                    var insertTasks = new List<InsertTask>();
                    int errorCount = 0;
                    var errorRows = new List<int>();

                    for (int i = 1; i <= lastRow - 1; i++)
                    {
                        int sheetRow = i + 1;
                        string expr = (exprVals[i, 1] ?? "").ToString().Trim();
                        if (string.IsNullOrEmpty(expr))
                        {
                            qtyOut[i - 1, 0] = "";
                            continue;
                        }

                        CalcResponse res = await EvalExprAsync(client, expr);
                        if (res == null || !res.Ok)
                        {
                            qtyOut[i - 1, 0] = "#ERR";
                            errorCount++;
                            errorRows.Add(sheetRow);
                            continue;
                        }

                        List<OutputDetail> details = BuildOutputDetails(res);
                        if (details != null && details.Count > 0)
                        {
                            OutputDetail primary = PickPrimary(details, res.PrimaryKey);
                            qtyOut[i - 1, 0] = primary.Value;

                            if (unitOut != null && !string.IsNullOrEmpty(primary.Unit))
                            {
                                unitOut[i - 1, 0] = primary.Unit;
                            }

                            var nonPrimary = new List<OutputDetail>();
                            foreach (var d in details)
                            {
                                if (!string.Equals(d.Key, primary.Key, StringComparison.OrdinalIgnoreCase))
                                {
                                    nonPrimary.Add(d);
                                }
                            }
                            if (nonPrimary.Count > 0)
                            {
                                string compName = !string.IsNullOrEmpty(res.ComponentName) ? res.ComponentName : ParseComponentName(expr);
                                insertTasks.Add(new InsertTask(sheetRow, compName, nonPrimary, expr));
                            }
                            continue;
                        }

                        if (res.Value.HasValue)
                        {
                            qtyOut[i - 1, 0] = res.Value.Value;
                        }
                        else
                        {
                            qtyOut[i - 1, 0] = "#ERR";
                            errorCount++;
                            errorRows.Add(sheetRow);
                        }

                    }

                    Excel.Range qtyRange = ws.Range[ws.Cells[2, qtyCol], ws.Cells[lastRow, qtyCol]];
                    qtyRange.Value2 = qtyOut;

                    if (unitOut != null)
                    {
                        Excel.Range unitRange = ws.Range[ws.Cells[2, unitCol], ws.Cells[lastRow, unitCol]];
                        unitRange.Value2 = unitOut;
                    }

                    if (insertTasks.Count > 0)
                    {
                        insertTasks.Sort((a, b) => b.RowNumber.CompareTo(a.RowNumber));
                        foreach (var task in insertTasks)
                        {
                            InsertOutputRows(ws, task, exprCol, qtyCol, rowTypeCol, nameCol, unitCol);
                        }
                    }

                    if (errorCount > 0)
                    {
                        string msg = BuildErrorSummary(errorCount, errorRows);
                        System.Windows.Forms.MessageBox.Show(msg, "ExcelSmartCost");
                    }
                }
            }
            finally
            {
                app.ScreenUpdating = prevSU;
                app.Calculation = prevCalc;
            }
        }

        private static CalcResponse EvalError(string message)
        {
            return new CalcResponse { Ok = false, Error = message ?? "Error" };
        }

        private static async Task<CalcResponse> EvalExprAsync(HttpClient client, string expr)
        {
            try
            {
                var req = new { expr = expr, variables = new Dictionary<string, double>(), precision = 6 };
                var payload = new StringContent(JsonConvert.SerializeObject(req), Encoding.UTF8, "application/json");
                var resp = await client.PostAsync("/api/calc/eval", payload);
                var json = await resp.Content.ReadAsStringAsync();
                CalcResponse res = null;
                try
                {
                    res = JsonConvert.DeserializeObject<CalcResponse>(json);
                }
                catch
                {
                    return EvalError("Invalid response");
                }
                if (!resp.IsSuccessStatusCode)
                {
                    return EvalError(res != null && !string.IsNullOrEmpty(res.Error) ? res.Error : $"HTTP {resp.StatusCode}");
                }
                return res ?? EvalError("Empty response");
            }
            catch (Exception ex)
            {
                return EvalError(ex.Message);
            }
        }

        private static List<OutputDetail> BuildOutputDetails(CalcResponse res)
        {
            if (res == null) return null;
            if (res.OutputsDetail != null && res.OutputsDetail.Count > 0)
            {
                return res.OutputsDetail;
            }
            if (res.Outputs == null || res.Outputs.Count == 0)
            {
                return null;
            }
            var keys = new List<string>(res.Outputs.Keys);
            keys.Sort(StringComparer.OrdinalIgnoreCase);
            var list = new List<OutputDetail>(keys.Count);
            foreach (var k in keys)
            {
                list.Add(new OutputDetail
                {
                    Key = k,
                    Name = k,
                    Unit = null,
                    Value = res.Outputs[k],
                    Primary = !string.IsNullOrEmpty(res.PrimaryKey) &&
                              string.Equals(k, res.PrimaryKey, StringComparison.OrdinalIgnoreCase),
                });
            }
            return list;
        }

        private static OutputDetail PickPrimary(List<OutputDetail> details, string primaryKey)
        {
            if (details == null || details.Count == 0) return null;
            foreach (var d in details)
            {
                if (d.Primary) return d;
            }
            if (!string.IsNullOrEmpty(primaryKey))
            {
                foreach (var d in details)
                {
                    if (string.Equals(d.Key, primaryKey, StringComparison.OrdinalIgnoreCase)) return d;
                }
            }
            return details[0];
        }

        private static string ParseComponentName(string expr)
        {
            if (string.IsNullOrEmpty(expr)) return "";
            int at = expr.IndexOf('@');
            if (at < 0) return "";
            int start = at + 1;
            int end = expr.IndexOf('(', start);
            if (end < 0) end = expr.Length;
            return expr.Substring(start, end - start).Trim();
        }

        private static string BuildErrorSummary(int errorCount, List<int> errorRows)
        {
            if (errorCount <= 0) return "Recalc completed with no errors.";
            int maxShow = 10;
            var shown = new List<string>();
            for (int i = 0; i < errorRows.Count && i < maxShow; i++)
            {
                shown.Add(errorRows[i].ToString());
            }
            string rowsPart = shown.Count > 0 ? string.Join(", ", shown) : "";
            string more = errorRows.Count > maxShow ? $" (+{errorRows.Count - maxShow} more)" : "";
            return $"Recalc completed with {errorCount} error(s). Rows: {rowsPart}{more}";
        }

        private static void InsertOutputRows(Excel.Worksheet ws, InsertTask task, int exprCol, int qtyCol, int rowTypeCol, int nameCol, int unitCol)
        {
            int count = task.Outputs.Count;
            if (count <= 0) return;
            int startRow = task.RowNumber + 1;
            Excel.Range insertRange = ws.Range[ws.Rows[startRow], ws.Rows[startRow + count - 1]];
            insertRange.Insert(Excel.XlInsertShiftDirection.xlShiftDown);

            var qtyVals = new object[count, 1];
            var nameVals = nameCol > 0 ? new object[count, 1] : null;
            var unitVals = unitCol > 0 ? new object[count, 1] : null;
            var rowTypeVals = rowTypeCol > 0 ? new object[count, 1] : null;
            var exprVals = exprCol > 0 ? new object[count, 1] : null;

            string compName = string.IsNullOrEmpty(task.ComponentName) ? "Component" : task.ComponentName;
            for (int i = 0; i < count; i++)
            {
                OutputDetail d = task.Outputs[i];
                qtyVals[i, 0] = d.Value;
                if (rowTypeVals != null) rowTypeVals[i, 0] = "COMP_OUTPUT";
                if (nameVals != null)
                {
                    string outName = string.IsNullOrEmpty(d.Name) ? d.Key : d.Name;
                    nameVals[i, 0] = $"【构件】{compName} - {outName}";
                }
                if (unitVals != null) unitVals[i, 0] = d.Unit ?? "";
                if (exprVals != null) exprVals[i, 0] = "";
            }

            if (rowTypeVals != null)
            {
                Excel.Range rtRange = ws.Range[ws.Cells[startRow, rowTypeCol], ws.Cells[startRow + count - 1, rowTypeCol]];
                rtRange.Value2 = rowTypeVals;
            }
            if (nameVals != null)
            {
                Excel.Range nameRange = ws.Range[ws.Cells[startRow, nameCol], ws.Cells[startRow + count - 1, nameCol]];
                nameRange.Value2 = nameVals;
            }
            if (unitVals != null)
            {
                Excel.Range unitRange = ws.Range[ws.Cells[startRow, unitCol], ws.Cells[startRow + count - 1, unitCol]];
                unitRange.Value2 = unitVals;
            }
            if (exprVals != null)
            {
                Excel.Range exprRange = ws.Range[ws.Cells[startRow, exprCol], ws.Cells[startRow + count - 1, exprCol]];
                exprRange.Value2 = exprVals;
            }
            Excel.Range qtyRange = ws.Range[ws.Cells[startRow, qtyCol], ws.Cells[startRow + count - 1, qtyCol]];
            qtyRange.Value2 = qtyVals;
        }

        private static void RemoveCompOutputRows(Excel.Worksheet ws, int rowTypeCol, int lastRow)
        {
            if (lastRow < 2) return;
            Excel.Range rtRange = ws.Range[ws.Cells[2, rowTypeCol], ws.Cells[lastRow, rowTypeCol]];
            object[,] rtVals = Ensure2D(rtRange.Value2, lastRow - 1, 1);
            var blocks = new List<Tuple<int, int>>();
            int start = -1;
            for (int i = 1; i <= lastRow - 1; i++)
            {
                string v = (rtVals[i, 1] ?? "").ToString().Trim();
                bool isComp = v.Equals("COMP_OUTPUT", StringComparison.OrdinalIgnoreCase);
                if (isComp && start == -1) start = i;
                if (!isComp && start != -1)
                {
                    blocks.Add(Tuple.Create(start, i - 1));
                    start = -1;
                }
            }
            if (start != -1)
            {
                blocks.Add(Tuple.Create(start, lastRow - 1));
            }
            for (int b = blocks.Count - 1; b >= 0; b--)
            {
                int s = blocks[b].Item1 + 1;
                int e = blocks[b].Item2 + 1;
                Excel.Range delRange = ws.Range[ws.Rows[s], ws.Rows[e]];
                delRange.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
            }
        }

        private static object[,] Ensure2D(object value, int rows, int cols)
        {
            if (value is object[,] arr) return arr;
            var wrapped = (object[,])Array.CreateInstance(typeof(object), new int[] { rows, cols }, new int[] { 1, 1 });
            wrapped[1, 1] = value;
            return wrapped;
        }

        private static Dictionary<string, int> ReadHeaders(Excel.Worksheet ws)
        {
            var dict = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int lastCol = ws.UsedRange.Columns.Count;
            for (int c = 1; c <= lastCol; c++)
            {
                var v = (ws.Cells[1, c] as Excel.Range).Value2;
                string s = v == null ? "" : v.ToString().Trim();
                if (!string.IsNullOrEmpty(s) && !dict.ContainsKey(s)) dict[s] = c;
            }
            return dict;
        }

        private sealed class InsertTask
        {
            public int RowNumber { get; }
            public string ComponentName { get; }
            public List<OutputDetail> Outputs { get; }
            public string Expr { get; }

            public InsertTask(int rowNumber, string componentName, List<OutputDetail> outputs, string expr)
            {
                RowNumber = rowNumber;
                ComponentName = componentName;
                Outputs = outputs ?? new List<OutputDetail>();
                Expr = expr;
            }
        }

        private sealed class CalcResponse
        {
            [JsonProperty("ok")]
            public bool Ok { get; set; }

            [JsonProperty("value")]
            public double? Value { get; set; }

            [JsonProperty("error")]
            public string Error { get; set; }

            [JsonProperty("outputs")]
            public Dictionary<string, double> Outputs { get; set; }

            [JsonProperty("primary_key")]
            public string PrimaryKey { get; set; }

            [JsonProperty("outputs_detail")]
            public List<OutputDetail> OutputsDetail { get; set; }

            [JsonProperty("component_name")]
            public string ComponentName { get; set; }
        }

        private sealed class OutputDetail
        {
            [JsonProperty("key")]
            public string Key { get; set; }

            [JsonProperty("name")]
            public string Name { get; set; }

            [JsonProperty("unit")]
            public string Unit { get; set; }

            [JsonProperty("value")]
            public double Value { get; set; }

            [JsonProperty("primary")]
            public bool Primary { get; set; }
        }
    }
}
