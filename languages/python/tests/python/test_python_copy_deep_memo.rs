use super::helpers::run_python;

#[test]
fn test_python_copy_deepcopy_basic() {
    let src = r#"
import copy
orig = {'a': [1, 2], 'b': {'x': 1}}
cloned = copy.deepcopy(orig)
cloned['a'].append(3)
cloned['b']['x'] = 9
print(orig['a'])
print(orig['b']['x'])
print(cloned['b']['x'])
"#;
    assert_eq!(run_python(src), vec!["[1, 2]", "1", "9"]);
}

#[test]
fn test_python_copy_deepcopy_keeps_independence() {
    let src = r#"
import copy
class Node:
    def __init__(self, value):
        self.value = value

n1 = Node(1)
n2 = copy.deepcopy(n1)
n2.value = 7
print(n1.value)
print(n2.value)
"#;
    assert_eq!(run_python(src), vec!["1", "7"]);
}
