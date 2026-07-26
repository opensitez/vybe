use super::helpers::run_python;

#[test]
fn test_python_bisect_left_right_boundaries() {
    let src = r#"
import bisect
values = [1, 3, 3, 5, 7, 9]
print(bisect.bisect_left(values, 3))
print(bisect.bisect_right(values, 3))
"#;
    assert_eq!(run_python(src), vec!["1", "3"]);
}

#[test]
fn test_python_bisect_insort_keeps_sorted() {
    let src = r#"
import bisect
values = [1, 4, 5]
bisect.insort(values, 3)
bisect.insort_left(values, 5)
print(values)
"#;
    assert_eq!(run_python(src), vec!["[1, 3, 4, 5, 5]"]);
}
