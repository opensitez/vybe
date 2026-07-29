use super::helpers::run_python;

#[test]
fn test_python_deque_pop_rotate() {
    let src = r#"
from collections import deque
q = deque([1, 2, 3, 4])
q.rotate(1)
q.appendleft(0)
print(list(q))
print(q.pop())
"#;
    assert_eq!(run_python(src), vec!["[0, 4, 1, 2, 3]", "3"]);
}

#[test]
fn test_python_deque_extend_left_capped() {
    let src = r#"
from collections import deque
q = deque([1, 2], maxlen=3)
q.extendleft([3, 4])
print(list(q))
"#;
    assert_eq!(run_python(src), vec!["[4, 3, 1]"]);
}
