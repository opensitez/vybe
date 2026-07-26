// Python os.environ — get, set, delete, getenv, keys/values
use super::helpers::run_python;

#[test]
fn test_environ_get_existing() {
    let script = r#"
import os
# PATH should always exist on Unix/Mac
print(isinstance(os.environ.get('PATH'), str))
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_environ_get_missing_default() {
    let script = r#"
import os
val = os.environ.get('__NO_SUCH_VAR_XYZ__', 'default')
print(val)
"#;
    assert_eq!(run_python(script), vec!["default"]);
}

#[test]
fn test_environ_set_and_get() {
    let script = r#"
import os
os.environ['_TEST_VAR_VYBE'] = 'hello'
print(os.environ['_TEST_VAR_VYBE'])
print(os.getenv('_TEST_VAR_VYBE'))
del os.environ['_TEST_VAR_VYBE']
"#;
    assert_eq!(run_python(script), vec!["hello", "hello"]);
}

#[test]
fn test_environ_delete_missing_raises() {
    let script = r#"
import os
try:
    del os.environ['__NO_SUCH_VAR_VYBE__']
    print("no_error")
except KeyError:
    print("KeyError")
"#;
    assert_eq!(run_python(script), vec!["KeyError"]);
}

#[test]
fn test_environ_getenv_none_default() {
    let script = r#"
import os
val = os.getenv('__MISSING_XYZ__')
print(val)
"#;
    assert_eq!(run_python(script), vec!["None"]);
}

#[test]
fn test_environ_keys_values_type() {
    let script = r#"
import os
print(type(os.environ.keys()).__name__)
print(len(os.environ.keys()) > 0)
"#;
    assert_eq!(run_python(script), vec!["KeysView", "True"]);
}

#[test]
fn test_environ_contains() {
    let script = r#"
import os
os.environ['_VYBE_TEST_VAR'] = 'x'
print('_VYBE_TEST_VAR' in os.environ)
del os.environ['_VYBE_TEST_VAR']
print('_VYBE_TEST_VAR' in os.environ)
"#;
    assert_eq!(run_python(script), vec!["True", "False"]);
}
