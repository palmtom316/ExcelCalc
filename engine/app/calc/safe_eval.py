from __future__ import annotations
import ast
import math
from typing import Dict, Callable

ALLOWED_NAMES = {"pi": math.pi, "PI": math.pi}
ALLOWED_FUNCS: Dict[str, Callable[..., float]] = {
    "sqrt": math.sqrt,
    "sq": lambda x: x * x,
    "pow": pow,
}

class SafeEvalError(Exception):
    pass

def _preprocess(expr: str) -> str:
    return expr.replace("^", "**")

ALLOWED_BINOPS = (ast.Add, ast.Sub, ast.Mult, ast.Div, ast.Pow)
ALLOWED_UNARYOPS = (ast.UAdd, ast.USub)

def safe_eval(expr: str, variables: Dict[str, float] | None = None) -> float:
    variables = variables or {}
    expr2 = _preprocess(expr.strip())
    try:
        node = ast.parse(expr2, mode="eval")
    except SyntaxError as e:
        raise SafeEvalError(f"Syntax error: {e}") from e

    def _eval(n: ast.AST) -> float:
        if isinstance(n, ast.Expression):
            return _eval(n.body)

        if isinstance(n, ast.Constant):
            if isinstance(n.value, (int, float)):
                return float(n.value)
            raise SafeEvalError("Only numeric constants are allowed")

        if isinstance(n, ast.BinOp):
            if not isinstance(n.op, ALLOWED_BINOPS):
                raise SafeEvalError("Operator not allowed")
            l = _eval(n.left); r = _eval(n.right)
            if isinstance(n.op, ast.Add): return l + r
            if isinstance(n.op, ast.Sub): return l - r
            if isinstance(n.op, ast.Mult): return l * r
            if isinstance(n.op, ast.Div):
                if r == 0: raise SafeEvalError("Division by zero")
                return l / r
            if isinstance(n.op, ast.Pow): return l ** r
            raise SafeEvalError("Operator not allowed")

        if isinstance(n, ast.UnaryOp):
            if not isinstance(n.op, ALLOWED_UNARYOPS):
                raise SafeEvalError("Unary operator not allowed")
            v = _eval(n.operand)
            return +v if isinstance(n.op, ast.UAdd) else -v

        if isinstance(n, ast.Name):
            if n.id in variables: return float(variables[n.id])
            if n.id in ALLOWED_NAMES: return float(ALLOWED_NAMES[n.id])
            raise SafeEvalError(f"Unknown variable: {n.id}")

        if isinstance(n, ast.Call):
            if not isinstance(n.func, ast.Name):
                raise SafeEvalError("Only simple function calls allowed")
            fname = n.func.id
            if fname not in ALLOWED_FUNCS:
                raise SafeEvalError(f"Function not allowed: {fname}")
            args = [_eval(a) for a in n.args]
            try:
                return float(ALLOWED_FUNCS[fname](*args))
            except ValueError as e:
                raise SafeEvalError(str(e)) from e

        raise SafeEvalError(f"Expression element not allowed: {type(n).__name__}")

    return _eval(node)
