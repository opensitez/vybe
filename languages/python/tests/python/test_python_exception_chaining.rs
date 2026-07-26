// Python exception chaining — raise from, __cause__, __context__, suppress
use super::helpers::run_python;

#[test]
fn test_raise_from_sets_cause() {
    let script = r#"
try:
    try:
        raise ValueError("original")
    except ValueError as e:
        raise RuntimeError("wrapped") from e
except RuntimeError as ex:
    print(ex)
    print(type(ex.__cause__).__name__)
    print(str(ex.__cause__))
"#;
    assert_eq!(run_python(script), vec!["wrapped", "ValueError", "original"]);
}

#[test]
fn test_implicit_chaining_context() {
    let script = r#"
try:
    try:
        raise ValueError("first")
    except ValueError:
        raise RuntimeError("second")
except RuntimeError as e:
    print(type(e.__context__).__name__)
"#;
    assert_eq!(run_python(script), vec!["ValueError"]);
}

#[test]
fn test_raise_from_none_suppresses_context() {
    let script = r#"
try:
    try:
        raise ValueError("original")
    except ValueError:
        raise RuntimeError("clean") from None
except RuntimeError as e:
    print(e.__cause__)
    print(e.__suppress_context__)
"#;
    assert_eq!(run_python(script), vec!["None", "True"]);
}

#[test]
fn test_contextlib_suppress() {
    let script = r#"
from contextlib import suppress

result = []
with suppress(ValueError):
    result.append(1)
    raise ValueError("ignored")
    result.append(2)  # never reached

result.append(3)
print(result)
"#;
    assert_eq!(run_python(script), vec!["[1, 3]"]);
}

#[test]
fn test_exception_chain_message() {
    let script = r#"
def process():
    try:
        1 / 0
    except ZeroDivisionError as e:
        raise PermissionError("cannot process") from e

try:
    process()
except PermissionError as e:
    print(str(e))
    print(type(e.__cause__).__name__)
"#;
    assert_eq!(run_python(script), vec!["cannot process", "ZeroDivisionError"]);
}

#[test]
fn test_nested_exception_groups() {
    let script = r#"
errors = []
for i in range(3):
    try:
        if i % 2 == 0:
            raise ValueError(f"even {i}")
    except ValueError as e:
        errors.append(str(e))

print(errors)
"#;
    assert_eq!(run_python(script), vec!["['even 0', 'even 2']"]);
}
