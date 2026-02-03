import pytest
from app.calc.safe_eval import safe_eval, SafeEvalError

def test_pi():
    assert abs(safe_eval("pi") - 3.141592653589793) < 1e-12

def test_sq_sqrt_pow():
    assert safe_eval("sq(0.3)") == pytest.approx(0.09)
    assert safe_eval("sqrt(9)") == 3.0
    assert safe_eval("2^3") == 8.0
    assert safe_eval("pow(2,3)") == 8.0

def test_basic_math():
    assert safe_eval("2*3+5") == 11.0
    assert safe_eval("(3-0.1)*20") == pytest.approx(58.0)

@pytest.mark.parametrize("expr",[
    '__import__("os")',
    'open("x","w")',
    '[].__class__',
    '().__class__.__mro__',
    '(lambda x:x)(1)',
])
def test_malicious(expr):
    with pytest.raises(SafeEvalError):
        safe_eval(expr)
