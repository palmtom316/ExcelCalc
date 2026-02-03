import pytest
from app.calc.component_engine import parse_component_call, ComponentDef, evaluate_component

def test_parse_component():
    name, params = parse_component_call("@Beam(width=0.3,height=0.5,len=10)")
    assert name == "Beam"
    assert params["width"] == pytest.approx(0.3)

def test_eval_component_multi_output_dependencies():
    comp = ComponentDef(
        name="X",
        params=[{"key":"L","required":True},{"key":"W","required":True},{"key":"H","required":True},{"key":"k","default":1.0}],
        outputs=[{"key":"a","primary":True},{"key":"b","primary":False}],
        formulas={"a":"L*W*H*k","b":"a-1"}
    )
    out = evaluate_component(comp, {"L":2,"W":3,"H":4})
    assert out["a"] == pytest.approx(24.0)
    assert out["b"] == pytest.approx(23.0)
