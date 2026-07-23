use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: sys + traceback — tracing, profiling, traceback formatting, intern, sys attributes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_traceback_format_exc() {
    let src = r#"
import traceback

try:
    1 / 0
except ZeroDivisionError:
    tb = traceback.format_exc()
    print("ZeroDivisionError" in tb)
    print("Traceback" in tb)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_traceback_extract_tb() {
    let src = r#"
import traceback, sys

def func_a():
    func_b()

def func_b():
    raise RuntimeError("error in b")

try:
    func_a()
except RuntimeError as e:
    tb = e.__traceback__
    frames = traceback.extract_tb(tb)
    func_names = [f.name for f in frames]
    print("func_a" in func_names)
    print("func_b" in func_names)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_intern_strings() {
    let src = r#"
import sys

s1 = sys.intern("dynamic_string_" + "123")
s2 = sys.intern("dynamic_string_" + "123")
print(s1 is s2)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_sys_settrace_line_events() {
    let src = r#"
import sys

events = []

def trace_calls(frame, event, arg):
    if event == "line":
        events.append(frame.f_lineno)
    return trace_calls

def target():
    a = 1
    b = 2
    c = a + b

sys.settrace(trace_calls)
target()
sys.settrace(None)

print(len(events) >= 3)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_sys_setprofile_function_calls() {
    let src = r#"
import sys

calls = []

def profile_func(frame, event, arg):
    if event == "call":
        calls.append(frame.f_code.co_name)

def helper():
    pass

def main():
    helper()

sys.setprofile(profile_func)
main()
sys.setprofile(None)

print("main" in calls)
print("helper" in calls)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_flags_and_dont_write_bytecode() {
    let src = r#"
import sys

print(isinstance(sys.dont_write_bytecode, bool))
print(isinstance(sys.flags, tuple) or hasattr(sys.flags, 'debug'))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_traceback_format_exception_only() {
    let src = r#"
import traceback

lines = traceback.format_exception_only(ValueError, ValueError("invalid value"))
formatted = "".join(lines).strip()
print(formatted)
"#;
    assert_eq!(run_python(src), vec!["ValueError: invalid value"]);
}

#[test]
fn test_py_sys_getsizeof() {
    let src = r#"
import sys

print(sys.getsizeof(1) > 0)
print(sys.getsizeof("hello") > 0)
print(sys.getsizeof([]) > 0)
print(sys.getsizeof({}) > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_sys_byteorder_and_encoding() {
    let src = r#"
import sys

print(sys.byteorder in ("little", "big"))
print(isinstance(sys.getdefaultencoding(), str))
print(sys.getdefaultencoding())
"#;
    assert_eq!(run_python(src), vec!["True", "True", "utf-8"]);
}

#[test]
fn test_py_traceback_print_stack() {
    let src = r#"
import traceback, io

buf = io.StringIO()
traceback.print_stack(file=buf)
output = buf.getvalue()
print("print_stack" in output or "test_py" in output)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
