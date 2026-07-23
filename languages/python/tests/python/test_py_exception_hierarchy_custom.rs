use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Exception Hierarchy & Custom Exceptions — BaseException, Exception, TypeError, ValueError, KeyError, AttributeError
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_exception_hierarchy_isinstance() {
    let src = r#"
e = KeyError("missing")
print(isinstance(e, KeyError))
print(isinstance(e, LookupError))
print(isinstance(e, Exception))
print(isinstance(e, BaseException))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_type_error_vs_value_error() {
    let src = r#"
try:
    len(42)
except TypeError as e:
    print("TypeError caught")

try:
    int("invalid_number")
except ValueError as e:
    print("ValueError caught")
"#;
    assert_eq!(
        run_python(src),
        vec!["TypeError caught", "ValueError caught"]
    );
}

#[test]
fn test_py_key_error_vs_index_error_lookup_base() {
    let src = r#"
def lookup(container, key):
    try:
        return container[key]
    except LookupError as e:
        return f"LookupError: {type(e).__name__}"

print(lookup({"a": 1}, "b"))
print(lookup([1, 2], 5))
"#;
    assert_eq!(
        run_python(src),
        vec!["LookupError: KeyError", "LookupError: IndexError"]
    );
}

#[test]
fn test_py_attribute_error_missing_attr() {
    let src = r#"
class Object: pass

obj = Object()
try:
    obj.missing_attribute()
except AttributeError as e:
    print("AttributeError caught")
"#;
    assert_eq!(run_python(src), vec!["AttributeError caught"]);
}

#[test]
fn test_py_custom_domain_exception_tree() {
    let src = r#"
class DomainError(Exception): pass
class UserError(DomainError): pass
class UserNotFoundError(UserError): pass

try:
    raise UserNotFoundError("User 123 not found")
except DomainError as e:
    print(f"Domain error: {type(e).__name__} - {e}")
"#;
    assert_eq!(
        run_python(src),
        vec!["Domain error: UserNotFoundError - User 123 not found"]
    );
}

#[test]
fn test_py_zero_division_error_arithmetic_base() {
    let src = r#"
try:
    1 / 0
except ArithmeticError as e:
    print(f"Arithmetic error: {type(e).__name__}")
"#;
    assert_eq!(run_python(src), vec!["Arithmetic error: ZeroDivisionError"]);
}

#[test]
fn test_py_exception_args_tuple() {
    let src = r#"
e = ValueError("arg1", "arg2", 3)
print(e.args)
print(e.args[0])
print(e.args[2])
"#;
    assert_eq!(run_python(src), vec!["('arg1', 'arg2', 3)", "arg1", "3"]);
}

#[test]
fn test_py_stop_iteration_exception_control_flow() {
    let src = r#"
it = iter([1])
print(next(it))
try:
    next(it)
except StopIteration:
    print("StopIteration caught")
"#;
    assert_eq!(run_python(src), vec!["1", "StopIteration caught"]);
}

#[test]
fn test_py_assertion_error_with_msg() {
    let src = r#"
try:
    assert 1 == 2, "1 does not equal 2"
except AssertionError as e:
    print(f"AssertionError: {e}")
"#;
    assert_eq!(run_python(src), vec!["AssertionError: 1 does not equal 2"]);
}

#[test]
fn test_py_recursion_error_max_depth() {
    let src = r#"
import sys
sys.setrecursionlimit(100)

def infinite_recursion():
    return infinite_recursion()

try:
    infinite_recursion()
except RecursionError:
    print("RecursionError caught")

sys.setrecursionlimit(1000)  # restore
"#;
    assert_eq!(run_python(src), vec!["RecursionError caught"]);
}
