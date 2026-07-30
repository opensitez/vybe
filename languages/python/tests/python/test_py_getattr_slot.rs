//! `__getattr__` — the attribute-miss interceptor, dispatched by protocol slot.
//!
//! Python `__getattr__`, PHP `__get` and JS Proxy get are one role
//! (`ProtocolSlot::GetAttr`). Both frontends already bind it; these cover the
//! shared dispatch that resolves it, so the role works without any site naming
//! a language's spelling.

use crate::helpers::run_python;

/// The plain case: a missing attribute reaches `__getattr__`.
#[test]
fn getattr_intercepts_a_missing_attribute() {
    let src = r#"
class Flex:
    def __getattr__(self, name):
        return f"default_{name}"

f = Flex()
print(f.missing)
"#;
    assert_eq!(run_python(src), vec!["default_missing"]);
}

/// An attribute that IS present must NOT reach `__getattr__` — Python calls it
/// only after normal lookup fails. This is the half a substitution would break.
#[test]
fn getattr_is_not_consulted_when_the_attribute_exists() {
    let src = r#"
class Flex:
    def __init__(self):
        self.real = "actual"

    def __getattr__(self, name):
        return "intercepted"

f = Flex()
print(f.real)
"#;
    assert_eq!(run_python(src), vec!["actual"]);
}

/// The interceptor receives the attribute NAME it was asked for.
#[test]
fn getattr_receives_the_attribute_name() {
    let src = r#"
class Recorder:
    def __getattr__(self, name):
        return name.upper()

r = Recorder()
print(r.alpha)
print(r.beta)
"#;
    assert_eq!(run_python(src), vec!["ALPHA", "BETA"]);
}

/// A class WITHOUT the role keeps ordinary behaviour — the probe must not
/// change what a plain read does.
#[test]
fn class_without_getattr_is_unaffected() {
    let src = r#"
class Plain:
    def __init__(self):
        self.here = 1

class Flex:
    def __getattr__(self, name):
        return "x"

p = Plain()
print(p.here)
"#;
    assert_eq!(run_python(src), vec!["1"]);
}

/// A miss on a class WITHOUT the role must raise `AttributeError` — Python
/// never surfaces the underlying storage as a `KeyError`.
#[test]
fn plain_miss_raises_attribute_error() {
    let src = r#"
class Plain:
    def __init__(self):
        self.here = 1

p = Plain()
try:
    print(p.nope)
except AttributeError:
    print("AttributeError")
except KeyError:
    print("KeyError")
"#;
    assert_eq!(run_python(src), vec!["AttributeError"]);
}
