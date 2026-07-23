use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Ctypes & C Data Structures — c_int, c_char_p, Structure, Union, POINTER, sizeof, cast, create_string_buffer
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_ctypes_c_primitives_sizeof() {
    let src = r#"
import ctypes

i = ctypes.c_int(42)
print(i.value)
print(ctypes.sizeof(ctypes.c_int))
print(ctypes.sizeof(ctypes.c_double))
"#;
    assert_eq!(run_python(src), vec!["42", "4", "8"]);
}

#[test]
fn test_py_ctypes_struct_layout_and_fields() {
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
fn test_py_ctypes_c_array_creation() {
    let src = r#"
import ctypes

IntArray3 = ctypes.c_int * 3
arr = IntArray3(10, 20, 30)

print(len(arr))
print(arr[0], arr[2])
print(list(arr))
"#;
    assert_eq!(run_python(src), vec!["3", "10 30", "[10, 20, 30]"]);
}

#[test]
fn test_py_ctypes_pointer_dereferencing() {
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
fn test_py_ctypes_string_buffer_mutation() {
    let src = r#"
import ctypes

buf = ctypes.create_string_buffer(b"Hello", 10)
print(buf.value.decode())
buf.value = b"World"
print(buf.value.decode())
"#;
    assert_eq!(run_python(src), vec!["Hello", "World"]);
}

#[test]
fn test_py_ctypes_union_memory_sharing() {
    let src = r#"
import ctypes

class DataUnion(ctypes.Union):
    _fields_ = [("i", ctypes.c_int), ("f", ctypes.c_float)]

u = DataUnion()
u.i = 1065353216  # IEEE 754 for 1.0f
print(round(u.f, 2))
"#;
    assert_eq!(run_python(src), vec!["1.0"]);
}

#[test]
fn test_py_ctypes_cast_pointer_conversion() {
    let src = r#"
import ctypes

arr = (ctypes.c_char * 4)(b"X", b"Y", b"Z", b"\x00")
ptr = ctypes.cast(arr, ctypes.c_char_p)
print(ptr.value.decode())
"#;
    assert_eq!(run_python(src), vec!["XYZ"]);
}

#[test]
fn test_py_ctypes_nested_structures() {
    let src = r#"
import ctypes

class Point(ctypes.Structure):
    _fields_ = [("x", ctypes.c_int), ("y", ctypes.c_int)]

class Rectangle(ctypes.Structure):
    _fields_ = [("top_left", Point), ("bottom_right", Point)]

r = Rectangle(Point(0, 0), Point(100, 200))
print(r.top_left.x, r.top_left.y)
print(r.bottom_right.x, r.bottom_right.y)
"#;
    assert_eq!(run_python(src), vec!["0 0", "100 200"]);
}

#[test]
fn test_py_ctypes_byref_argument_passing() {
    let src = r#"
import ctypes

val = ctypes.c_int(42)
ref = ctypes.byref(val)
print(isinstance(ref, ctypes._CData) or type(ref).__name__ == "CArgObject")
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_ctypes_c_char_p_null_pointer() {
    let src = r#"
import ctypes

null_ptr = ctypes.c_char_p(None)
print(null_ptr.value is None)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}
