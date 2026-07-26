// Python slice objects — slice(), start/stop/step, indices(), extended slicing
use super::helpers::run_python;

#[test]
fn test_slice_object_creation() {
    let script = r#"
s = slice(1, 5)
print(s.start, s.stop, s.step)
"#;
    assert_eq!(run_python(script), vec!["1 5 None"]);
}

#[test]
fn test_slice_object_with_step() {
    let script = r#"
s = slice(0, 10, 2)
lst = list(range(10))
print(lst[s])
"#;
    assert_eq!(run_python(script), vec!["[0, 2, 4, 6, 8]"]);
}

#[test]
fn test_slice_indices_method() {
    let script = r#"
s = slice(1, 10, 2)
indices = s.indices(7)
print(indices)
"#;
    assert_eq!(run_python(script), vec!["(1, 7, 2)"]);
}

#[test]
fn test_slice_negative_step() {
    let script = r#"
lst = [1, 2, 3, 4, 5]
print(lst[::-1])
print(lst[4:1:-1])
"#;
    assert_eq!(run_python(script), vec!["[5, 4, 3, 2, 1]", "[5, 4, 3]"]);
}

#[test]
fn test_slice_open_ended() {
    let script = r#"
lst = list(range(10))
print(lst[:5])
print(lst[5:])
print(lst[:])
"#;
    assert_eq!(run_python(script), vec!["[0, 1, 2, 3, 4]", "[5, 6, 7, 8, 9]", "[0, 1, 2, 3, 4, 5, 6, 7, 8, 9]"]);
}

#[test]
fn test_slice_out_of_bounds_safe() {
    let script = r#"
lst = [1, 2, 3]
print(lst[1:100])
print(lst[-100:2])
"#;
    assert_eq!(run_python(script), vec!["[2, 3]", "[1, 2]"]);
}

#[test]
fn test_custom_getitem_with_slice() {
    let script = r#"
class SparseList:
    def __init__(self, data):
        self.data = data
    def __getitem__(self, key):
        if isinstance(key, slice):
            return self.data[key.start:key.stop]
        return self.data[key]

sl = SparseList([10, 20, 30, 40, 50])
print(sl[1:4])
print(sl[2])
"#;
    assert_eq!(run_python(script), vec!["[20, 30, 40]", "30"]);
}
