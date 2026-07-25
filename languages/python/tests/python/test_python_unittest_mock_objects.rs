use super::helpers::run_python;

// unittest.mock — Mock, MagicMock, patch, patch.object, patch.dict, ANY, PropertyMock, side_effect, return_value, assert_called_with, call

#[test]
fn test_mock_return_value_and_called_with() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock(return_value=42)
res = m(1, 2, key="val")
print(res)
m.assert_called_with(1, 2, key="val")
print("assert_called_with passed")
"#);
    assert_eq!(out, vec!["42", "assert_called_with passed"]);
}

#[test]
fn test_mock_side_effect_exception() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock(side_effect=KeyError("missing_key"))
try:
    m()
except KeyError as e:
    print("KeyError:", str(e))
"#);
    assert_eq!(out, vec!["KeyError: 'missing_key'"]);
}

#[test]
fn test_mock_side_effect_iterable_sequence() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock(side_effect=[10, 20, 30])
print(m(), m(), m())
"#);
    assert_eq!(out, vec!["10 20 30"]);
}

#[test]
fn test_mock_side_effect_callable_function() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock(side_effect=lambda x: x * 2)
print(m(5), m(10))
"#);
    assert_eq!(out, vec!["10 20"]);
}

#[test]
fn test_mock_magic_mock_dunder_methods() {
    let out = run_python(r#"
from unittest.mock import MagicMock
m = MagicMock()
m.__str__.return_value = "custom_str"
m.__len__.return_value = 5
print(str(m))
print(len(m))
"#);
    assert_eq!(out, vec!["custom_str", "5"]);
}

#[test]
fn test_mock_patch_decorator() {
    let out = run_python(r#"
from unittest.mock import patch, Mock
import os

@patch("os.getcwd", return_value="/mocked/path")
def run_test(mock_getcwd):
    print(os.getcwd())
    mock_getcwd.assert_called_once()

run_test()
"#);
    assert_eq!(out, vec!["/mocked/path"]);
}

#[test]
fn test_mock_patch_context_manager() {
    let out = run_python(r#"
from unittest.mock import patch
import sys

with patch("sys.platform", "mocked_os"):
    print(sys.platform)

print(sys.platform != "mocked_os")
"#);
    assert_eq!(out, vec!["mocked_os", "True"]);
}

#[test]
fn test_mock_patch_object() {
    let out = run_python(r#"
from unittest.mock import patch

class Calculator:
    def add(self, a, b):
        return a + b

calc = Calculator()
with patch.object(calc, "add", return_value=999):
    print(calc.add(2, 3))

print(calc.add(2, 3))
"#);
    assert_eq!(out, vec!["999", "5"]);
}

#[test]
fn test_mock_patch_dict() {
    let out = run_python(r#"
from unittest.mock import patch
import os

env_override = {"MY_VAR": "mocked_val"}
with patch.dict(os.environ, env_override, clear=False):
    print(os.environ.get("MY_VAR"))

print(os.environ.get("MY_VAR") is None)
"#);
    assert_eq!(out, vec!["mocked_val", "True"]);
}

#[test]
fn test_mock_any_wildcard_matcher() {
    let out = run_python(r#"
from unittest.mock import Mock, ANY
m = Mock()
m(42, "string", [1, 2])
m.assert_called_with(42, ANY, ANY)
print("ANY match passed")
"#);
    assert_eq!(out, vec!["ANY match passed"]);
}

#[test]
fn test_mock_property_mock() {
    let out = run_python(r#"
from unittest.mock import MagicMock, PropertyMock

class Foo:
    @property
    def val(self): return 100

foo = Foo()
with patch.object(Foo, "val", new_callable=PropertyMock) as mock_prop:
    mock_prop.return_value = 500
    print(foo.val)

print(foo.val)
"#);
    assert_eq!(out, vec!["500", "100"]);
}

#[test]
fn test_mock_call_args_and_call_count() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock()
m(1, a="foo")
m(2, a="bar")
print(m.call_count)
print(m.call_args)
"#);
    assert_eq!(out, vec!["2", "call(2, a='bar')"]);
}

#[test]
fn test_mock_call_args_list() {
    let out = run_python(r#"
from unittest.mock import Mock, call
m = Mock()
m("first")
m("second")
print(m.call_args_list == [call("first"), call("second")])
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_mock_assert_called_once_with() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock()
m("only_once")
m.assert_called_once_with("only_once")
print("assert_called_once_with passed")
"#);
    assert_eq!(out, vec!["assert_called_once_with passed"]);
}

#[test]
fn test_mock_assert_has_calls() {
    let out = run_python(r#"
from unittest.mock import Mock, call
m = Mock()
m(1)
m(2)
m(3)
m.assert_has_calls([call(1), call(2)], any_order=False)
print("assert_has_calls passed")
"#);
    assert_eq!(out, vec!["assert_has_calls passed"]);
}

#[test]
fn test_mock_reset_mock() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock(return_value=1)
m(100)
print(m.called)
m.reset_mock()
print(m.called)
print(m.call_count)
"#);
    assert_eq!(out, vec!["True", "False", "0"]);
}

#[test]
fn test_mock_spec_restricts_attributes() {
    let out = run_python(r#"
from unittest.mock import Mock

class Greeter:
    def hello(self): pass

m = Mock(spec=Greeter)
m.hello()  # allowed
try:
    m.non_existent_method()  # raises AttributeError
except AttributeError:
    print("AttributeError")
"#);
    assert_eq!(out, vec!["AttributeError"]);
}

#[test]
fn test_mock_seal_prevents_new_attributes() {
    let out = run_python(r#"
from unittest.mock import Mock, seal, sys
if sys.version_info >= (3, 8):
    m = Mock()
    m.existing = 1
    seal(m)
    try:
        m.new_attr = 2
    except AttributeError:
        print("AttributeError")
else:
    print("AttributeError")
"#);
    assert_eq!(out, vec!["AttributeError"]);
}

#[test]
fn test_mock_attach_mock() {
    let out = run_python(r#"
from unittest.mock import Mock
parent = Mock()
child = Mock()
parent.attach_mock(child, "child_func")
child("child_call")
print(parent.mock_calls)
"#);
    assert_eq!(out, vec!["[call.child_func('child_call')]"]);
}

#[test]
fn test_mock_assert_never_called() {
    let out = run_python(r#"
from unittest.mock import Mock
m = Mock()
m.assert_never_called()
print("assert_never_called passed")
"#);
    assert_eq!(out, vec!["assert_never_called passed"]);
}
