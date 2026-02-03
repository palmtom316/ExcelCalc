你是资深 Windows Excel 插件工程师（VSTO/C#）+ Python 后端工程师。请基于“路线A”实现 ExcelSmartCost 工程量计算系统 MVP。

# 目标
在 Excel 明细表内输入表达式（数学/构件），点击 Ribbon “重算当前表”，自动写回 Quantity；构件多输出（方案B）需展开 COMP_OUTPUT 子行。

# 约束
- Windows only
- 离线可用（不依赖外网）
- 数学表达式必须安全解析（禁止 eval/exec）
- 支持：pi/PI、sqrt、sq、pow、幂运算 ^
- 构件：@ComponentName(p=v,...)，多输出，允许 outputs 相互引用

# 工程交付
请输出可运行 Python Engine + 可调试 VSTO Add-in 骨架：
- engine/: FastAPI + SQLite + safe_eval + component_engine + importers + pytest
- add-in/: Ribbon + EngineClient + QuantityRecalc（批量读写）
必须实现：
- 批处理关闭 ScreenUpdating，Calculation=manual
- 错误写回 #ERR，结束汇总弹窗
- 多输出：一次插入 N 行并批量写入（禁止循环插入单行）

# 输出要求
1) 文件树
2) 逐文件完整内容
3) 运行步骤与测试命令
现在开始生成代码。
