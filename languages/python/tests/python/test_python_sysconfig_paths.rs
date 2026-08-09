use super::helpers::run_python;

// sysconfig — get_path, get_paths, get_config_var, get_scheme_names, get_platform

#[test]
fn test_sysconfig_get_scheme_names_list() {
    let out = run_python(
        r#"
import sysconfig
names = sysconfig.get_scheme_names()
print(isinstance(names, tuple))
print(len(names) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_get_path_names() {
    let out = run_python(
        r#"
import sysconfig
names = sysconfig.get_path_names()
for name in ("stdlib", "platstdlib", "platlib", "purelib", "include", "scripts", "data"):
    print(name in names)
"#,
    );
    assert_eq!(
        out,
        vec!["True", "True", "True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_sysconfig_get_path_stdlib() {
    let out = run_python(
        r#"
import sysconfig, os
path = sysconfig.get_path("stdlib")
print(isinstance(path, str))
print(os.path.exists(path))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_get_path_scripts_is_string() {
    let out = run_python(
        r#"
import sysconfig
path = sysconfig.get_path("scripts")
print(isinstance(path, str))
print(len(path) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_get_paths_returns_dict() {
    let out = run_python(
        r#"
import sysconfig
paths = sysconfig.get_paths()
print(isinstance(paths, dict))
print("stdlib" in paths)
print("scripts" in paths)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_sysconfig_get_config_var_prefix() {
    let out = run_python(
        r#"
import sysconfig, sys
prefix = sysconfig.get_config_var("prefix")
print(isinstance(prefix, str))
print(len(prefix) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_get_config_var_unknown_returns_none() {
    let out = run_python(
        r#"
import sysconfig
val = sysconfig.get_config_var("NONEXISTENT_VAR_XYZ123")
print(val)
"#,
    );
    assert_eq!(out, vec!["None"]);
}

#[test]
fn test_sysconfig_get_config_var_version() {
    let out = run_python(
        r#"
import sysconfig, sys
version = sysconfig.get_config_var("py_version")
major_minor = f"{sys.version_info.major}.{sys.version_info.minor}"
print(version.startswith(major_minor))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_get_config_vars_returns_dict() {
    let out = run_python(
        r#"
import sysconfig
cfg = sysconfig.get_config_vars()
print(isinstance(cfg, dict))
print(len(cfg) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_get_platform_is_string() {
    let out = run_python(
        r#"
import sysconfig
p = sysconfig.get_platform()
print(isinstance(p, str))
print(len(p) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_get_platform_contains_arch() {
    let out = run_python(
        r#"
import sysconfig
p = sysconfig.get_platform()
# Platform string like "linux-x86_64" or "macosx-12.0-arm64"
print("-" in p)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_get_python_version() {
    let out = run_python(
        r#"
import sysconfig, sys
v = sysconfig.get_python_version()
print(v == f"{sys.version_info.major}.{sys.version_info.minor}")
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_default_scheme_in_names() {
    let out = run_python(
        r#"
import sysconfig
default = sysconfig.get_default_scheme()
print(default in sysconfig.get_scheme_names())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_get_path_with_explicit_scheme() {
    let out = run_python(
        r#"
import sysconfig
scheme = sysconfig.get_default_scheme()
path = sysconfig.get_path("stdlib", scheme)
print(isinstance(path, str))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_parse_config_h() {
    let out = run_python(
        r#"
import sysconfig, io
sample = "/* config.h */\n#define PY_MAJOR_VERSION 3\n#define Py_DEBUG 0\n"
result = {}
sysconfig.parse_config_h(io.StringIO(sample), result)
print(result.get("PY_MAJOR_VERSION"))
print(result.get("Py_DEBUG"))
"#,
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn test_sysconfig_get_config_var_py_version_short() {
    let out = run_python(
        r#"
import sysconfig
v = sysconfig.get_config_var("py_version_short")
parts = v.split(".")
print(len(parts) == 2)
print(parts[0].isdigit())
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_sysconfig_is_python_build() {
    let out = run_python(
        r#"
import sysconfig
result = sysconfig.is_python_build()
print(isinstance(result, bool))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_get_path_include_is_string() {
    let out = run_python(
        r#"
import sysconfig
path = sysconfig.get_path("include")
print(isinstance(path, str))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_paths_all_strings() {
    let out = run_python(
        r#"
import sysconfig
paths = sysconfig.get_paths()
print(all(isinstance(v, str) for v in paths.values()))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_sysconfig_get_config_vars_subset() {
    let out = run_python(
        r#"
import sysconfig
result = sysconfig.get_config_vars("py_version", "prefix")
print(len(result) == 2)
print(all(isinstance(v, str) for v in result))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}
