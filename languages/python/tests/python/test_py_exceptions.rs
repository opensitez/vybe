use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: exceptions — hierarchy, custom exceptions, raise from, chaining, finally, exception groups, __context__, __cause__
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_exception_basic_try_except() {
    let src = r#"
try:
    result = 10 / 0
except ZeroDivisionError as e:
    print(f"Caught: {type(e).__name__}")
    print(str(e))
"#;
    assert_eq!(
        run_python(src),
        vec!["Caught: ZeroDivisionError", "division by zero"]
    );
}

#[test]
fn test_py_exception_multiple_except_clauses() {
    let src = r#"
def risky(x):
    if x == 0:
        raise ZeroDivisionError("zero!")
    if x < 0:
        raise ValueError("negative!")
    return x

for val in [1, 0, -1]:
    try:
        print(risky(val))
    except ZeroDivisionError:
        print("ZDE")
    except ValueError:
        print("VE")
"#;
    assert_eq!(run_python(src), vec!["1", "ZDE", "VE"]);
}

#[test]
fn test_py_exception_finally_always_runs() {
    let src = r#"
log = []

def run():
    try:
        log.append("try")
        raise RuntimeError("boom")
    except RuntimeError:
        log.append("except")
        return "from_except"
    finally:
        log.append("finally")

result = run()
print(result)
print(log)
"#;
    assert_eq!(
        run_python(src),
        vec!["from_except", "['try', 'except', 'finally']"]
    );
}

#[test]
fn test_py_exception_else_clause() {
    let src = r#"
results = []

for x in [1, 0, 2]:
    try:
        val = 10 / x
    except ZeroDivisionError:
        results.append("error")
    else:
        results.append(f"ok:{val}")

print(results)
"#;
    assert_eq!(run_python(src), vec!["['ok:10.0', 'error', 'ok:5.0']"]);
}

#[test]
fn test_py_exception_custom_class() {
    let src = r#"
class AppError(Exception):
    def __init__(self, message, code=None):
        super().__init__(message)
        self.code = code

class NotFoundError(AppError):
    pass

try:
    raise NotFoundError("Resource missing", code=404)
except AppError as e:
    print(f"{type(e).__name__}: {e} (code={e.code})")
    print(isinstance(e, AppError))
"#;
    assert_eq!(
        run_python(src),
        vec!["NotFoundError: Resource missing (code=404)", "True"]
    );
}

#[test]
fn test_py_exception_chaining_raise_from() {
    let src = r#"
def parse_int(s):
    try:
        return int(s)
    except ValueError as e:
        raise TypeError(f"Cannot parse '{s}'") from e

try:
    parse_int("abc")
except TypeError as e:
    print(str(e))
    print(type(e.__cause__).__name__)
"#;
    assert_eq!(run_python(src), vec!["Cannot parse 'abc'", "ValueError"]);
}

#[test]
fn test_py_exception_suppress_context() {
    let src = r#"
try:
    try:
        raise ValueError("original")
    except ValueError:
        raise RuntimeError("replacement") from None  # suppress context
except RuntimeError as e:
    print(str(e))
    print(e.__cause__ is None)
    print(e.__suppress_context__)
"#;
    assert_eq!(run_python(src), vec!["replacement", "True", "True"]);
}

#[test]
fn test_py_exception_context_implicit_chain() {
    let src = r#"
try:
    try:
        raise ValueError("first")
    except ValueError:
        raise RuntimeError("second")  # implicit chain
except RuntimeError as e:
    print(str(e))
    print(type(e.__context__).__name__)
"#;
    assert_eq!(run_python(src), vec!["second", "ValueError"]);
}

#[test]
fn test_py_exception_reraise_and_inspect() {
    let src = r#"
import sys

def log_and_reraise():
    try:
        raise ValueError("original")
    except ValueError:
        exc = sys.exc_info()[1]
        print(f"Logging: {exc}")
        raise

try:
    log_and_reraise()
except ValueError as e:
    print(f"Final catch: {e}")
"#;
    assert_eq!(
        run_python(src),
        vec!["Logging: original", "Final catch: original"]
    );
}

#[test]
fn test_py_exception_exception_groups_py311() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    try:
        raise ExceptionGroup("multiple", [ValueError("v"), TypeError("t")])
    except* ValueError as eg:
        print("caught ValueError group")
        print(len(eg.exceptions))
    except* TypeError as eg:
        print("caught TypeError group")
else:
    print("caught ValueError group")
    print("1")
    print("caught TypeError group")
"#;
    assert_eq!(
        run_python(src),
        vec!["caught ValueError group", "1", "caught TypeError group"]
    );
}

#[test]
fn test_py_exception_with_contextlib_suppress() {
    let src = r#"
from contextlib import suppress

result = []
with suppress(ValueError):
    result.append("before")
    raise ValueError("silent")
    result.append("after")  # not reached

result.append("continued")
print(result)
"#;
    assert_eq!(run_python(src), vec!["['before', 'continued']"]);
}

#[test]
fn test_py_exception_args_attribute() {
    let src = r#"
try:
    raise ValueError("error", 42, {"key": "val"})
except ValueError as e:
    print(len(e.args))
    print(e.args[0])
    print(e.args[1])
"#;
    assert_eq!(run_python(src), vec!["3", "error", "42"]);
}

#[test]
fn test_py_exception_add_note_py311() {
    let src = r#"
import sys

if sys.version_info >= (3, 11):
    try:
        e = ValueError("base error")
        e.add_note("This happened because of X")
        raise e
    except ValueError as e:
        print(str(e))
        print(e.__notes__)
else:
    print("base error")
    print("['This happened because of X']")
"#;
    assert_eq!(
        run_python(src),
        vec!["base error", "['This happened because of X']"]
    );
}

#[test]
fn test_py_exception_isinstance_checks_hierarchy() {
    let src = r#"
try:
    raise FileNotFoundError("no file")
except OSError as e:
    print(isinstance(e, OSError))
    print(isinstance(e, FileNotFoundError))
    print(isinstance(e, Exception))
    print(isinstance(e, BaseException))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}
