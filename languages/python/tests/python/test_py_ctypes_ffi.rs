use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: ctypes — C data types, structures, pointers, sizeof, function prototypes
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_ctypes_primitive_types() {
    let src = r#"
import ctypes

i = ctypes.c_int(42)
print(i.value)

f = ctypes.c_double(3.14)
print(round(f.value, 2))

s = ctypes.c_char_p(b"hello c")
print(s.value.decode())
"#;
    assert_eq!(run_python(src), vec!["42", "3.14", "hello c"]);
}

#[test]
fn test_py_ctypes_sizeof_and_alignment() {
    let src = r#"
import ctypes

print(ctypes.sizeof(ctypes.c_int))
print(ctypes.sizeof(ctypes.c_double))
print(ctypes.sizeof(ctypes.c_char))
print(ctypes.sizeof(ctypes.c_void_p))
"#;
    assert_eq!(run_python(src), vec!["4", "8", "1", "8"]); // 64-bit arch
}

#[test]
fn test_py_ctypes_struct_definition() {
    let src = r#"
import ctypes

class Point(ctypes.Structure):
    _fields_ = [("x", ctypes.c_int), ("y", ctypes.c_int)]

p = Point(10, 20)
print(p.x, p.y)
print(ctypes.sizeof(Point))
"#;
    assert_eq!(run_python(src), vec!["10 20", "8"]);
}

#[test]
fn test_py_ctypes_array_creation() {
    let src = r#"
import ctypes

IntArray5 = ctypes.c_int * 5
arr = IntArray5(1, 2, 3, 4, 5)

print(len(arr))
print(arr[0], arr[4])
print(list(arr))
"#;
    assert_eq!(run_python(src), vec!["5", "1 5", "[1, 2, 3, 4, 5]"]);
}

#[test]
fn test_py_ctypes_pointer_and_dereference() {
    let src = r#"
import ctypes

val = ctypes.c_int(100)
ptr = ctypes.pointer(val)

print(ptr.contents.value)
ptr.contents.value = 200
print(val.value)
"#;
    assert_eq!(run_python(src), vec!["100", "200"]);
}

#[test]
fn test_py_ctypes_cast() {
    let src = r#"
import ctypes

arr = (ctypes.c_char * 4)(b"A", b"B", b"C", b"\x00")
ptr = ctypes.cast(arr, ctypes.c_char_p)
print(ptr.value.decode())
"#;
    assert_eq!(run_python(src), vec!["ABC"]);
}

#[test]
fn test_py_ctypes_union() {
    let src = r#"
import ctypes

class Data(ctypes.Union):
    _fields_ = [("i", ctypes.c_int), ("f", ctypes.c_float)]

d = Data()
d.i = 1065353216  # binary representation of 1.0f in IEEE 754
print(round(d.f, 2))
"#;
    assert_eq!(run_python(src), vec!["1.0"]);
}

#[test]
fn test_py_ctypes_byref() {
    let src = r#"
import ctypes

val = ctypes.c_int(42)
ref = ctypes.byref(val)
print(isinstance(ref, ctypes._CData) or type(ref).__name__ == "CArgObject")
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_ctypes_string_buffer() {
    let src = r#"
import ctypes

buf = ctypes.create_string_buffer(b"Hello", 10)
print(buf.value.decode())
print(len(buf.raw))
buf.value = b"World"
print(buf.value.decode())
"#;
    assert_eq!(run_python(src), vec!["Hello", "10", "World"]);
}

#[test]
fn test_py_ctypes_libc_math_abs() {
    let src = r#"
import ctypes, sys

if sys.platform.startswith("darwin") or sys.platform.startswith("linux"):
    libc = ctypes.CDLL(None)
    abs_func = libc.abs
    abs_func.argtypes = [ctypes.c_int]
    abs_func.restype = ctypes.c_int
    print(abs_func(-42))
else:
    print(42)
"#;
    assert_eq!(run_python(src), vec!["42"]);
}
