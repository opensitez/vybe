use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: import system — __import__, importlib, dynamic imports, __all__, __name__, relative imports, sys.modules
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_import_standard_module() {
    let src = r#"
import math
from math import pi, sqrt
from math import ceil as ceiling

print(round(pi, 5))
print(sqrt(16))
print(ceiling(4.1))
"#;
    assert_eq!(run_python(src), vec!["3.14159", "4.0", "5"]);
}

#[test]
fn test_py_import_module_attributes() {
    let src = r#"
import os

print(isinstance(os.__name__, str))
print(os.__name__)
print(isinstance(os.__file__, str))
print(os.__spec__ is not None)
"#;
    assert_eq!(run_python(src), vec!["True", "os", "True", "True"]);
}

#[test]
fn test_py_import_sys_modules_caching() {
    let src = r#"
import sys
import math
import math as math2

print(math is math2)
print("math" in sys.modules)
print(sys.modules["math"] is math)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_import_importlib_dynamic() {
    let src = r#"
import importlib

json = importlib.import_module("json")
result = json.dumps({"key": "value"})
print(result)
print(json.__name__)
"#;
    assert_eq!(run_python(src), vec![r#"{"key": "value"}"#, "json"]);
}

#[test]
fn test_py_import_dunder_name_in_script() {
    let src = r#"
# When run as a module (not as __main__), __name__ == module name
print(__name__)
"#;
    // When run by our harness, __name__ is typically "__main__" or the module name
    let result = run_python(src);
    assert!(!result.is_empty());
    assert!(!result[0].is_empty());
}

#[test]
fn test_py_import_star_with_all() {
    let src = r#"
import types, sys

# Simulate a module with __all__
mod = types.ModuleType("fake_mod")
mod.__all__ = ["public_func"]
mod.public_func = lambda: "public"
mod._private = "private"
sys.modules["fake_mod"] = mod

from fake_mod import *
print(public_func())

# _private should NOT be imported via *
print("_private" not in dir())
"#;
    assert_eq!(run_python(src), vec!["public", "True"]);
}

#[test]
fn test_py_import_reload() {
    let src = r#"
import importlib, sys, types

# Create a simple module dynamically
mod = types.ModuleType("reload_test")
mod.value = 42
sys.modules["reload_test"] = mod

import reload_test
print(reload_test.value)

mod.value = 99
importlib.reload(reload_test)
print(reload_test.value)
"#;
    assert_eq!(run_python(src), vec!["42", "99"]);
}

#[test]
fn test_py_import_module_spec_and_loader() {
    let src = r#"
import importlib.util
import json

spec = importlib.util.find_spec("json")
print(spec is not None)
print(spec.name)
print(spec.loader is not None)
"#;
    assert_eq!(run_python(src), vec!["True", "json", "True"]);
}

#[test]
fn test_py_import_pkgutil_iter_modules() {
    let src = r#"
import pkgutil, sys

# Check that standard library modules are discoverable
names = [m.name for m in pkgutil.iter_modules() if m.name in ("json", "math", "os")]
print(sorted(set(names)))
"#;
    assert_eq!(run_python(src), vec!["['json', 'math', 'os']"]);
}

#[test]
fn test_py_import_lazy_with_importlib() {
    let src = r#"
import importlib

module_name = "collections"
mod = importlib.import_module(module_name)
Counter = getattr(mod, "Counter")
c = Counter("aabbc")
print(c.most_common(1))
"#;
    assert_eq!(run_python(src), vec!["[('a', 2)]"]);
}

#[test]
fn test_py_import_namespace_and_packages() {
    let src = r#"
import email.mime.text
import xml.etree.ElementTree as ET

# Basic accessibility checks
print(email.mime.text.__name__)
print(ET.__name__)
"#;
    assert_eq!(
        run_python(src),
        vec!["email.mime.text", "xml.etree.ElementTree"]
    );
}
