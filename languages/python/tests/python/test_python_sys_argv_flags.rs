// Python sys module — argv, version, path, modules, exc_info, getrefcount
use super::helpers::run_python;

#[test]
fn test_sys_version_info() {
    let script = r#"
import sys
print(sys.version_info.major >= 3)
print(sys.version_info.minor >= 0)
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_sys_modules_contains_builtin() {
    let script = r#"
import sys
import os
print('os' in sys.modules)
print('sys' in sys.modules)
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_sys_path_is_list() {
    let script = r#"
import sys
print(isinstance(sys.path, list))
"#;
    assert_eq!(run_python(script), vec!["True"]);
}

#[test]
fn test_sys_platform() {
    let script = r#"
import sys
print(isinstance(sys.platform, str))
print(len(sys.platform) > 0)
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_sys_maxsize() {
    let script = r#"
import sys
print(sys.maxsize > 0)
print(isinstance(sys.maxsize, int))
"#;
    assert_eq!(run_python(script), vec!["True", "True"]);
}

#[test]
fn test_sys_exc_info_outside_handler() {
    let script = r#"
import sys
exc_type, exc_val, exc_tb = sys.exc_info()
print(exc_type)
"#;
    assert_eq!(run_python(script), vec!["None"]);
}

#[test]
fn test_sys_exc_info_in_handler() {
    let script = r#"
import sys
try:
    raise ValueError("test")
except:
    exc_type, exc_val, exc_tb = sys.exc_info()
    print(exc_type.__name__)
    print(str(exc_val))
"#;
    assert_eq!(run_python(script), vec!["ValueError", "test"]);
}

#[test]
fn test_sys_byteorder() {
    let script = r#"
import sys
print(sys.byteorder in ('little', 'big'))
"#;
    assert_eq!(run_python(script), vec!["True"]);
}
