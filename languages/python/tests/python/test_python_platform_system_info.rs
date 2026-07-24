use super::helpers::run_python;

// platform — uname, python_version, machine, architecture, system, node, release

#[test]
fn test_platform_system_is_string() {
    let out = run_python(r#"
import platform
s = platform.system()
print(s in ["Linux", "Darwin", "Windows", "FreeBSD", "NetBSD", "OpenBSD"])
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_node_is_nonempty_string() {
    let out = run_python(r#"
import platform
print(len(platform.node()) > 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_python_version_tuple() {
    let out = run_python(r#"
import platform
t = platform.python_version_tuple()
print(len(t))
print(t[0].isdigit())
print(t[1].isdigit())
"#);
    assert_eq!(out, vec!["3", "True", "True"]);
}

#[test]
fn test_platform_python_version_string_format() {
    let out = run_python(r#"
import platform
v = platform.python_version()
parts = v.split(".")
print(len(parts) >= 2)
print(parts[0].isdigit())
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_platform_machine_is_string() {
    let out = run_python(r#"
import platform
m = platform.machine()
print(isinstance(m, str))
print(len(m) > 0)
"#);
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_platform_processor_is_string() {
    let out = run_python(r#"
import platform
p = platform.processor()
print(isinstance(p, str))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_architecture_bits() {
    let out = run_python(r#"
import platform
bits, linkage = platform.architecture()
print(bits in ["32bit", "64bit"])
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_uname_namedtuple_fields() {
    let out = run_python(r#"
import platform
u = platform.uname()
print(hasattr(u, "system"))
print(hasattr(u, "node"))
print(hasattr(u, "release"))
print(hasattr(u, "version"))
print(hasattr(u, "machine"))
"#);
    assert_eq!(out, vec!["True", "True", "True", "True", "True"]);
}

#[test]
fn test_platform_uname_system_matches_system() {
    let out = run_python(r#"
import platform
print(platform.uname().system == platform.system())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_python_implementation() {
    let out = run_python(r#"
import platform
impl = platform.python_implementation()
print(impl in ["CPython", "PyPy", "Jython", "IronPython"])
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_python_compiler_nonempty() {
    let out = run_python(r#"
import platform
c = platform.python_compiler()
print(len(c) > 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_python_build_returns_tuple() {
    let out = run_python(r#"
import platform
build = platform.python_build()
print(len(build) == 2)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_release_is_string() {
    let out = run_python(r#"
import platform
r = platform.release()
print(isinstance(r, str))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_version_is_string() {
    let out = run_python(r#"
import platform
v = platform.version()
print(isinstance(v, str))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_python_revision_is_string() {
    let out = run_python(r#"
import platform
r = platform.python_revision()
print(isinstance(r, str))
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_uname_node_matches_node() {
    let out = run_python(r#"
import platform
print(platform.uname().node == platform.node())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_uname_machine_matches_machine() {
    let out = run_python(r#"
import platform
print(platform.uname().machine == platform.machine())
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_platform_string_nonempty() {
    let out = run_python(r#"
import platform
print(len(platform.platform()) > 0)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_python_version_tuple_major_is_3() {
    let out = run_python(r#"
import platform
major, _, _ = platform.python_version_tuple()
print(int(major) >= 3)
"#);
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_platform_architecture_executable_present() {
    let out = run_python(r#"
import platform, sys
bits, _ = platform.architecture(sys.executable)
print(bits in ["32bit", "64bit"])
"#);
    assert_eq!(out, vec!["True"]);
}
