use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Exception Handling Flow — try-except-else-finally, chaining, reraise, custom exceptions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_try_except_else_finally_execution_order() {
    let src = r#"
events = []

def run(fail):
    events.clear()
    try:
        events.append("try")
        if fail:
            raise ValueError("boom")
    except ValueError:
        events.append("except")
    else:
        events.append("else")
    finally:
        events.append("finally")
    return list(events)

print(run(False))
print(run(True))
"#;
    assert_eq!(
        run_python(src),
        vec!["['try', 'else', 'finally']", "['try', 'except', 'finally']"]
    );
}

#[test]
fn test_py_finally_return_override() {
    let src = r#"
def func():
    try:
        return "try"
    finally:
        return "finally"  # overrides try return

print(func())
"#;
    assert_eq!(run_python(src), vec!["finally"]);
}

#[test]
fn test_py_explicit_exception_chaining_from() {
    let src = r#"
class DatabaseError(Exception): pass

def query():
    try:
        int("invalid")
    except ValueError as cause:
        raise DatabaseError("Query failed") from cause

try:
    query()
except DatabaseError as e:
    print(e)
    print(type(e.__cause__).__name__)
"#;
    assert_eq!(run_python(src), vec!["Query failed", "ValueError"]);
}

#[test]
fn test_py_implicit_exception_chaining_context() {
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
        raise RuntimeError("Clean error") from None
except RuntimeError as e:
    print(e.__cause__ is None)
    print(e.__suppress_context__)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_custom_exception_attributes() {
    let src = r#"
class ValidationError(Exception):
    def __init__(self, field, message, code):
        super().__init__(f"{field}: {message}")
        self.field = field
        self.message = message
        self.code = code

try:
    raise ValidationError("email", "invalid format", 400)
except ValidationError as e:
    print(e.field)
    print(e.message)
    print(e.code)
    print(str(e))
"#;
    assert_eq!(
        run_python(src),
        vec!["email", "invalid format", "400", "email: invalid format"]
    );
}

#[test]
fn test_py_multiple_exception_tuple_matching() {
    let src = r#"
def parse(val):
    try:
        if val == "zero":
            1 / 0
        elif val == "int":
            int("abc")
        elif val == "key":
            {}[val]
    except (ZeroDivisionError, ValueError, KeyError) as e:
        print(f"Caught {type(e).__name__}")

parse("zero")
parse("int")
parse("key")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Caught ZeroDivisionError",
            "Caught ValueError",
            "Caught KeyError"
        ]
    );
}

#[test]
fn test_py_reraise_exception() {
    let src = r#"
log = []

def process():
    try:
        raise KeyError("missing_key")
    except KeyError as e:
        log.append("logged keyerror")
        raise

try:
    process()
except KeyError as e:
    log.append(f"re-caught {e}")

print(log)
"#;
    assert_eq!(
        run_python(src),
        vec!["['logged keyerror', \"re-caught 'missing_key'\"]"]
    );
}

#[test]
fn test_py_sys_exc_info_inspection() {
    let src = r#"
import sys

try:
    raise TypeError("bad type")
except TypeError:
    exc_type, exc_val, exc_tb = sys.exc_info()
    print(exc_type.__name__)
    print(str(exc_val))
"#;
    assert_eq!(run_python(src), vec!["TypeError", "bad type"]);
}

#[test]
fn test_py_exception_notes_py311() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    try:
        err = ValueError("bad value")
        err.add_note("Context info: field X")
        raise err
    except ValueError as e:
        print(e.__notes__[0])
else:
    print("Context info: field X")
"#;
    assert_eq!(run_python(src), vec!["Context info: field X"]);
}
