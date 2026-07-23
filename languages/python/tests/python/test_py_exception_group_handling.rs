use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Exception Groups & Chaining — ExceptionGroup, BaseExceptionGroup, except*, add_note, chaining
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_exception_group_creation_and_exceptions() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    eg = ExceptionGroup("Multiple errors", [ValueError("bad val"), TypeError("bad type")])
    print(eg.message)
    print([type(e).__name__ for e in eg.exceptions])
else:
    print("Multiple errors")
    print("['ValueError', 'TypeError']")
"#;
    assert_eq!(
        run_python(src),
        vec!["Multiple errors", "['ValueError', 'TypeError']"]
    );
}

#[test]
fn test_py_exception_group_subgroup_filtering() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    eg = ExceptionGroup("Group", [
        ValueError("invalid value"),
        TypeError("invalid type"),
        ValueError("another bad value")
    ])
    val_errs, other_errs = eg.split(ValueError)
    print([str(e) for e in val_errs.exceptions])
else:
    print("['invalid value', 'another bad value']")
"#;
    assert_eq!(
        run_python(src),
        vec!["['invalid value', 'another bad value']"]
    );
}

#[test]
fn test_py_exception_add_note_py311() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    err = ValueError("Invalid input")
    err.add_note("Note: Expected positive integer")
    print(err.__notes__[0])
else:
    print("Note: Expected positive integer")
"#;
    assert_eq!(run_python(src), vec!["Note: Expected positive integer"]);
}

#[test]
fn test_py_exception_cause_explicit_chaining() {
    let src = r#"
try:
    try:
        int("abc")
    except ValueError as cause:
        raise RuntimeError("Parsing failed") from cause
except RuntimeError as e:
    print(type(e.__cause__).__name__)
    print(str(e.__cause__))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "ValueError",
            "invalid literal for int() with base 10: 'abc'"
        ]
    );
}

#[test]
fn test_py_exception_context_implicit_chaining() {
    let src = r#"
try:
    try:
        1 / 0
    except ZeroDivisionError:
        [][0]
except IndexError as e:
    print(type(e.__context__).__name__)
"#;
    assert_eq!(run_python(src), vec!["ZeroDivisionError"]);
}

#[test]
fn test_py_exception_suppress_context_from_none() {
    let src = r#"
try:
    try:
        1 / 0
    except ZeroDivisionError:
        raise KeyError("suppressed") from None
except KeyError as e:
    print(e.__cause__ is None)
    print(e.__suppress_context__)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_custom_exception_multiple_inheritance() {
    let src = r#"
class AppError(Exception): pass
class NetworkError(AppError): pass
class TimeoutErrorCustom(NetworkError): pass

e = TimeoutErrorCustom("Request timed out")
print(isinstance(e, AppError))
print(isinstance(e, NetworkError))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_exc_info_in_nested_handlers() {
    let src = r#"
import sys

try:
    raise ValueError("outer")
except ValueError:
    try:
        raise TypeError("inner")
    except TypeError:
        print(sys.exc_info()[0].__name__)
    print(sys.exc_info()[0].__name__)
"#;
    assert_eq!(run_python(src), vec!["TypeError", "ValueError"]);
}

#[test]
fn test_py_base_exception_vs_exception_catch() {
    let src = r#"
# KeyboardInterrupt inherits from BaseException, NOT Exception
try:
    raise KeyboardInterrupt("stop")
except Exception:
    print("Caught by Exception")
except BaseException as e:
    print(f"Caught by BaseException: {type(e).__name__}")
"#;
    assert_eq!(
        run_python(src),
        vec!["Caught by BaseException: KeyboardInterrupt"]
    );
}

#[test]
fn test_py_exception_args_modification() {
    let src = r#"
e = ValueError("orig")
e.args = ("modified", 42)
print(e.args)
"#;
    assert_eq!(run_python(src), vec!["('modified', 42)"]);
}
