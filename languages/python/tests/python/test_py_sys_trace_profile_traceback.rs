use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Sys Trace, Profile & Traceback — sys.settrace, sys.setprofile, traceback.format_exc, extract_tb, sys.intern, getsizeof
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_traceback_format_exc_string_contains_error() {
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
fn test_py_traceback_extract_tb_frame_stack() {
    let src = r#"
import traceback

def alpha():
    beta()

def beta():
    raise RuntimeError("error in beta")

try:
    alpha()
except RuntimeError as e:
    frames = traceback.extract_tb(e.__traceback__)
    func_names = [f.name for f in frames]
    print("alpha" in func_names)
    print("beta" in func_names)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_intern_string_reference_equality() {
    let src = r#"
import sys

s1 = sys.intern("dynamic_str_" + "key")
s2 = sys.intern("dynamic_str_" + "key")
print(s1 is s2)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_sys_settrace_line_execution_events() {
    let src = r#"
import sys

events = []

def tracer(frame, event, arg):
    if event == "line":
        events.append(frame.f_lineno)
    return tracer

def target_func():
    a = 1
    b = 2
    return a + b

sys.settrace(tracer)
target_func()
sys.settrace(None)

print(len(events) >= 3)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_sys_setprofile_call_return_events() {
    let src = r#"
import sys

calls = []

def profiler(frame, event, arg):
    if event == "call":
        calls.append(frame.f_code.co_name)

def helper(): pass
def main(): helper()

sys.setprofile(profiler)
main()
sys.setprofile(None)

print("main" in calls)
print("helper" in calls)
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_getsizeof_memory_allocation() {
    let src = r#"
import sys

print(sys.getsizeof(0) > 0)
print(sys.getsizeof("string") > 0)
print(sys.getsizeof([]) > 0)
print(sys.getsizeof({}) > 0)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_py_traceback_format_exception_only_single_line() {
    let src = r#"
import traceback

lines = traceback.format_exception_only(ValueError, ValueError("invalid value"))
formatted = "".join(lines).strip()
print(formatted)
"#;
    assert_eq!(run_python(src), vec!["ValueError: invalid value"]);
}

#[test]
fn test_py_sys_flags_and_dont_write_bytecode_configuration() {
    let src = r#"
import sys

print(isinstance(sys.dont_write_bytecode, bool))
print(isinstance(sys.flags.optimize, int))
"#;
    assert_eq!(run_python(src), vec!["True", "True"]);
}

#[test]
fn test_py_sys_byteorder_and_encoding_properties() {
    let src = r#"
import sys

print(sys.byteorder in ("little", "big"))
print(sys.getdefaultencoding())
"#;
    assert_eq!(run_python(src), vec!["True", "utf-8"]);
}

#[test]
fn test_py_traceback_print_stack_output_capture() {
    let src = r#"
import traceback, io

buf = io.StringIO()
traceback.print_stack(file=buf)
out = buf.getvalue()
print("print_stack" in out or "test_py" in out)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
