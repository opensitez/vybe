use super::helpers::run_python;

// traceback — format_exc, format_exception, extract_tb, format_list, FrameSummary, StackSummary, TracebackException, print_exc

#[test]
fn test_traceback_format_exc_captures_current_exception() {
    let out = run_python(
        r#"
import traceback
try:
    1 / 0
except ZeroDivisionError:
    tb_str = traceback.format_exc()
    print("ZeroDivisionError: division by zero" in tb_str)
    print("Traceback (most recent call last):" in tb_str)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_traceback_format_exception_list() {
    let out = run_python(
        r#"
import traceback, sys
try:
    raise ValueError("invalid parameter")
except ValueError as exc:
    lines = traceback.format_exception(exc)
    full_text = "".join(lines)
    print("ValueError: invalid parameter" in full_text)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_extract_tb_returns_frame_summaries() {
    let out = run_python(
        r#"
import traceback, sys

def level2(): raise RuntimeError("fail")
def level1(): level2()

try:
    level1()
except RuntimeError as exc:
    tb = exc.__traceback__
    frames = traceback.extract_tb(tb)
    print(len(frames) >= 2)
    print(frames[-1].name)
"#,
    );
    assert_eq!(out, vec!["True", "level2"]);
}

#[test]
fn test_traceback_frame_summary_attributes() {
    let out = run_python(
        r#"
import traceback, sys

try:
    1 / 0
except ZeroDivisionError as exc:
    frames = traceback.extract_tb(exc.__traceback__)
    f = frames[-1]
    print(isinstance(f.filename, str))
    print(isinstance(f.lineno, int))
    print(isinstance(f.name, str))
    print(isinstance(f.line, str))
"#,
    );
    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn test_traceback_traceback_exception_from_exception() {
    let out = run_python(
        r#"
import traceback
try:
    raise KeyError("missing")
except KeyError as exc:
    te = traceback.TracebackException.from_exception(exc)
    print(te.exc_type.__name__)
    formatted = "".join(te.format())
    print("KeyError: 'missing'" in formatted)
"#,
    );
    assert_eq!(out, vec!["KeyError", "True"]);
}

#[test]
fn test_traceback_format_list_from_extracted_frames() {
    let out = run_python(
        r#"
import traceback
try:
    int("not_a_number")
except ValueError as exc:
    frames = traceback.extract_tb(exc.__traceback__)
    formatted_list = traceback.format_list(frames)
    print(len(formatted_list) == len(frames))
    print("line" in formatted_list[0].lower() or "file" in formatted_list[0].lower())
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_traceback_print_exc_to_string_stream() {
    let out = run_python(
        r#"
import traceback, io
buf = io.StringIO()
try:
    [][0]  # IndexError
except IndexError:
    traceback.print_exc(file=buf)

output = buf.getvalue()
print("IndexError: list index out of range" in output)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_walk_tb_generator() {
    let out = run_python(
        r#"
import traceback

def a(): b()
def b(): c()
def c(): raise TypeError("type_err")

try:
    a()
except TypeError as exc:
    frames = list(traceback.walk_tb(exc.__traceback__))
    print(len(frames) >= 3)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_walk_stack_generator() {
    let out = run_python(
        r#"
import traceback
frames = list(traceback.walk_stack(None))
print(len(frames) > 0)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_stack_summary_extract() {
    let out = run_python(
        r#"
import traceback
stack = traceback.StackSummary.extract(traceback.walk_stack(None))
print(len(stack) > 0)
print(isinstance(stack[0], traceback.FrameSummary))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_traceback_stack_summary_format() {
    let out = run_python(
        r#"
import traceback
stack = traceback.StackSummary.extract(traceback.walk_stack(None))
formatted = stack.format()
print(isinstance(formatted, list))
print(len(formatted) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_traceback_traceback_exception_notes_support() {
    let out = run_python(
        r#"
import traceback, sys
if sys.version_info >= (3, 11):
    try:
        e = ValueError("orig error")
        e.add_note("Note 1: check config")
        raise e
    except ValueError as exc:
        te = traceback.TracebackException.from_exception(exc)
        formatted = "".join(te.format())
        print("Note 1: check config" in formatted)
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_format_exception_only() {
    let out = run_python(
        r#"
import traceback
lines = traceback.format_exception_only(ValueError, ValueError("bad arg"))
text = "".join(lines).strip()
print(text)
"#,
    );
    assert_eq!(out, vec!["ValueError: bad arg"]);
}

#[test]
fn test_traceback_clear_frames() {
    let out = run_python(
        r#"
import traceback
try:
    1 / 0
except ZeroDivisionError as exc:
    tb = exc.__traceback__
    traceback.clear_frames(tb)
    print("frames cleared")
"#,
    );
    assert_eq!(out, vec!["frames cleared"]);
}

#[test]
fn test_traceback_format_stack_current_execution() {
    let out = run_python(
        r#"
import traceback
stack_lines = traceback.format_stack()
print(isinstance(stack_lines, list))
print(len(stack_lines) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_traceback_print_tb_to_stream() {
    let out = run_python(
        r#"
import traceback, io
buf = io.StringIO()
try:
    raise NameError("var not found")
except NameError as exc:
    traceback.print_tb(exc.__traceback__, file=buf)

print("line" in buf.getvalue().lower())
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_chained_exceptions_cause_formatting() {
    let out = run_python(
        r#"
import traceback
try:
    try:
        1 / 0
    except ZeroDivisionError as cause:
        raise RuntimeError("wrapper error") from cause
except RuntimeError as exc:
    te = traceback.TracebackException.from_exception(exc)
    formatted = "".join(te.format())
    print("The above exception was the direct cause" in formatted)
    print("ZeroDivisionError: division by zero" in formatted)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_traceback_implicit_chained_exceptions_context() {
    let out = run_python(
        r#"
import traceback
try:
    try:
        int("bad")
    except ValueError:
        [][0]
except IndexError as exc:
    te = traceback.TracebackException.from_exception(exc)
    formatted = "".join(te.format())
    print("During handling of the above exception" in formatted)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_traceback_exception_max_group_depth() {
    let out = run_python(
        r#"
import traceback, sys
if sys.version_info >= (3, 11):
    eg = ExceptionGroup("group", [ValueError(1), TypeError(2)])
    te = traceback.TracebackException.from_exception(eg)
    formatted = "".join(te.format())
    print("ExceptionGroup: group" in formatted)
else:
    print(True)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_traceback_extract_stack_current() {
    let out = run_python(
        r#"
import traceback
stack = traceback.extract_stack()
print(isinstance(stack, traceback.StackSummary))
print(len(stack) > 0)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}
