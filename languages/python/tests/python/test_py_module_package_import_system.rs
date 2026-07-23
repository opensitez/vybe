use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Module & Package Import System — __import__, importlib, sys.modules, __all__, dynamic imports
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dynamic_import_with_importlib() {
    let src = r#"
import importlib

math = importlib.import_module("math")
print(math.sqrt(25))
print(math.pi > 3.14)
"#;
    assert_eq!(run_python(src), vec!["5.0", "True"]);
}

#[test]
fn test_py_sys_modules_cache_inspection() {
    let src = r#"
import sys, json

print("json" in sys.modules)
print(sys.modules["json"] is json)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_builtin_import_function() {
    let src = r#"
re = __import__("re")
match = re.search(r"\d+", "abc123def")
print(match.group())
"#;
    assert_eq!(run_python(src), vec!["123"]);
}

#[test]
fn test_py_importlib_util_find_spec() {
    let src = r#"
import importlib.util

spec = importlib.util.find_spec("os")
print(spec is not None)
print(spec.name)
"#;
    assert_eq!(run_python(src), vec!["True", "os"]);
}

#[test]
fn test_py_fake_module_creation_in_sys_modules() {
    let src = r#"
import sys, types

mod = types.ModuleType("my_custom_module")
mod.custom_val = 42
mod.custom_func = lambda x: x * 2

sys.modules["my_custom_module"] = mod

import my_custom_module
print(my_custom_module.custom_val)
print(my_custom_module.custom_func(10))
"#;
    assert_eq!(run_python(src), vec!["42", "20"]);
}

#[test]
fn test_py_all_export_filtering() {
    let src = r#"
import sys, types

mod = types.ModuleType("exported_mod")
mod.__all__ = ["allowed_func"]
mod.allowed_func = lambda: "allowed"
mod.disallowed_func = lambda: "disallowed"

sys.modules["exported_mod"] = mod

from exported_mod import *
print(allowed_func())
print("disallowed_func" in globals())
"#;
    assert_eq!(run_python(src), vec!["allowed", "False"]);
}

#[test]
fn test_py_importlib_reload_module() {
    let src = r#"
import sys, types, importlib

mod = types.ModuleType("reloadable")
mod.version = 1
sys.modules["reloadable"] = mod

import reloadable
print(reloadable.version)

mod.version = 2
importlib.reload(reloadable)
print(reloadable.version)
"#;
    assert_eq!(run_python(src), vec!["1", "2"]);
}

#[test]
fn test_py_module_attributes_name_file_doc() {
    let src = r#"
import json

print(json.__name__)
print(isinstance(json.__doc__, str))
print(hasattr(json, "__file__"))
"#;
    assert_eq!(run_python(src), vec!["json", "True", "True"]);
}

#[test]
fn test_py_import_error_handling() {
    let src = r#"
try:
    import nonexistent_module_abc_xyz
except ImportError as e:
    print(f"ImportError: {e.name}")
"#;
    assert_eq!(
        run_python(src),
        vec!["ImportError: nonexistent_module_abc_xyz"]
    );
}

#[test]
fn test_py_from_import_as_alias() {
    let src = r#"
from math import sqrt as root, pi as PI

print(root(16))
print(round(PI, 2))
"#;
    assert_eq!(run_python(src), vec!["4.0", "3.14"]);
}
