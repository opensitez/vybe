// Python array module — typed arrays, append, extend, typecodes, frombytes
use super::helpers::run_python;

#[test]
fn test_array_creation_int() {
    let script = r#"
import array
a = array.array('i', [1, 2, 3, 4])
print(len(a))
print(a[0])
print(a[-1])
"#;
    assert_eq!(run_python(script), vec!["4", "1", "4"]);
}

#[test]
fn test_array_typecode() {
    let script = r#"
import array
a = array.array('d', [1.5, 2.5])
print(a.typecode)
"#;
    assert_eq!(run_python(script), vec!["d"]);
}

#[test]
fn test_array_append_extend() {
    let script = r#"
import array
a = array.array('i', [1, 2])
a.append(3)
a.extend([4, 5])
print(list(a))
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3, 4, 5]"]);
}

#[test]
fn test_array_pop_remove() {
    let script = r#"
import array
a = array.array('i', [10, 20, 30, 20])
a.remove(20)
print(list(a))
popped = a.pop()
print(popped)
"#;
    assert_eq!(run_python(script), vec!["[10, 30, 20]", "20"]);
}

#[test]
fn test_array_index() {
    let script = r#"
import array
a = array.array('i', [5, 10, 15, 10])
print(a.index(10))
"#;
    assert_eq!(run_python(script), vec!["1"]);
}

#[test]
fn test_array_count() {
    let script = r#"
import array
a = array.array('i', [1, 2, 1, 3, 1])
print(a.count(1))
"#;
    assert_eq!(run_python(script), vec!["3"]);
}

#[test]
fn test_array_reverse() {
    let script = r#"
import array
a = array.array('i', [1, 2, 3])
a.reverse()
print(list(a))
"#;
    assert_eq!(run_python(script), vec!["[3, 2, 1]"]);
}

#[test]
fn test_array_tobytes_frombytes() {
    let script = r#"
import array
a = array.array('B', [65, 66, 67])
b = a.tobytes()
print(b)
a2 = array.array('B')
a2.frombytes(b)
print(list(a2))
"#;
    assert_eq!(run_python(script), vec!["b'ABC'", "[65, 66, 67]"]);
}

#[test]
fn test_array_tolist() {
    let script = r#"
import array
a = array.array('f', [1.0, 2.0, 3.0])
lst = a.tolist()
print(type(lst).__name__)
print(len(lst))
"#;
    assert_eq!(run_python(script), vec!["list", "3"]);
}

#[test]
fn test_array_unsigned_byte() {
    let script = r#"
import array
a = array.array('B', [0, 127, 255])
print(a[0])
print(a[2])
"#;
    assert_eq!(run_python(script), vec!["0", "255"]);
}
