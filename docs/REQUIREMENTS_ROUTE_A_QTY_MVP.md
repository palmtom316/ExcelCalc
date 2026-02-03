# ExcelSmartCost（路线A）工程量计算系统 MVP 需求文档
版本：v1.0  
平台：Windows + Microsoft Excel（WPS 兼容加分但不要求专门版本）  
架构：VSTO/C# Add-in（控 Excel） + Python Engine（计算/入库/SQLite）  
离线：必须离线可用（不依赖外网）  

---

## 1. 目标与范围

### 1.1 本阶段目标（只做工程量计算闭环）
- 在 Excel 明细表中输入工程量表达式（数学表达式或构件调用）
- 一键重算：自动写回工程量结果（Quantity）
- 构件库（方案B，多输出）：构件调用可产出多个工程量项
  - primary 输出写回当前行 Quantity
  - 非 primary 输出以“子行（COMP_OUTPUT）”形式插入并写回

### 1.2 Out of Scope（后续版本）
- 清单库/定额库/企业价格库的检索、插入、组价、汇总（本阶段不做业务闭环）
- 报表/汇总表
- 在线服务、云端依赖

---

## 2. 用户工作流（MVP）

### 2.1 Excel 表内输入
用户在 `Calc_Expr` 列输入：
- 数学表达式：`(3-0.1)*20`、`pi*sq(2)/4`、`sqrt(9)`、`2^3`
- 构件调用：`@Beam(width=0.3,height=0.5,len=10)`

### 2.2 一键计算
用户点击 Ribbon：
- 【重算当前表】→ 批量读取 Calc_Expr → 调用 Python Engine → 批量写回 Quantity
- 若表达式为构件且有多输出：插入 N 行子行并写回各输出（方案B）

---

## 3. Excel 模板规范（MVP 必需）

活动 Sheet 第1行为表头，至少包含：

| Header | 必需 | 说明 |
|---|---|---|
| Calc_Expr | 是 | 表达式输入 |
| Quantity | 是 | 计算结果写回 |
| RowType | 否（建议） | `DATA` / `COMP_OUTPUT` |
| RowID | 否（建议） | 唯一ID（可由插件生成） |
| Name | 否（建议） | 名称/输出项名称 |
| Unit | 否（建议） | 单位 |

**缺表头处理：**
- 若缺少 `Calc_Expr` 或 `Quantity`：弹窗提示并退出，不做写入。

---

## 4. 表达式规范（重点）

### 4.1 数学表达式（Math）
支持：
- 运算符：`+ - * /`
- 幂：`^`（推荐）与 `pow(a,b)`（可同时支持）
- 常量：`pi` / `PI`
- 函数白名单：`sqrt(x)`（开方）、`sq(x)`（平方）、`pow(x,y)`
- 括号：`()`

**不支持/禁止：**
- 任何属性访问、下标、导入、调用系统函数、lambda 等

**错误规则：**
- 语法错误/除零/不在白名单：返回错误，Quantity 写 `#ERR`
- `sqrt(负数)`：返回错误，Quantity 写 `#ERR`

### 4.2 构件调用（Component，方案B）
语法：
`@ComponentName(p1=v1, p2=v2, ...)`

- `ComponentName`：字母/数字/下划线
- 参数值 `v`：数字或可由数学表达式计算（MVP 可先限制为数值/简单表达式）
- 构件输出：多个 outputs，按构件定义公式计算，允许 outputs 之间引用（按定义顺序）

**落地规则：**
- primary 输出：写回当前行 Quantity（Unit 同步写回，如有 Unit 列）
- 非 primary 输出：在当前行下插入 N 行子行
  - RowType = `COMP_OUTPUT`
  - Name = `【构件】{ComponentName} - {output_name}`
  - Quantity = output value
  - Unit = output unit
  - Calc_Expr 可写公式（可选，建议写，便于追溯）

---

## 5. 构件库（入库、参数、公式）

### 5.1 SQLite 表（lib_components_v2）
- name (PK)
- params_json：JSON array（参数定义）
- outputs_json：JSON array（输出定义，含 primary）
- formulas_json：JSON object（output_key -> formula）
- version, description

### 5.2 MVP 内置构件清单（先入库，公式可占位待图纸修订）
- 电力管沟 PowerCableTrench
- 管道包封 PipeEncasement
- 电力井 PowerCableManhole
- 电力顶管 PowerPipeJacking
- 电力顶管工作井 PowerJackingWorkingShaft
- 设备基础 EquipmentFoundation
- 环网柜基础 RingMainUnitFoundation

---

## 6. 系统接口（Python Engine）

### 6.1 鉴权
所有 `/api/*` 需 Header：
- `X-Token: <token>`

### 6.2 计算接口：POST /api/calc/eval
请求：
```json
{"expr":"pi*sq(2)/4","variables":{},"precision":6}
```

响应（构件）：
```json
{"ok":true,"value":1.5,"outputs":{"volume_m3":1.5},"primary_key":"volume_m3"}
```

### 6.3 入库导入：POST /api/import
请求：
```json
{"kind":"components","file_path":"C:/path/components.xlsx","sheet_name":"components","dedup_mode":"overwrite"}
```

---

## 7. Add-in（C#）实现要求（性能/稳定性）
- 批量读取/写回 Range
- 计算期间关闭 ScreenUpdating，Calculation=manual
- 多输出子行：一次插入 N 行（禁止循环插入单行）
- 错误写回 `#ERR`，结束弹窗汇总

---

## 8. 验收标准（AC）
- 支持 pi/sq/sqrt/^
- 支持 @Component 多输出展开子行
- 恶意表达式必须失败且无系统调用
- 批量性能可用（不卡死）
