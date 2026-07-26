use super::helpers::run_python;

#[test]
fn test_python_chainmap_lookup_and_mutation() {
    let src = r#"
from collections import ChainMap
base = {'a': 1}
over = {'b': 2}
cm = ChainMap(base, over)
print(cm['a'])
base['a'] = 10
print(cm['a'])
"#;
    assert_eq!(run_python(src), vec!["1", "10"]);
}

#[test]
fn test_python_chainmap_new_child_and_parents() {
    let src = r#"
from collections import ChainMap
cm = ChainMap({'a': 1}, {'b': 2})
child = cm.new_child({'a': 9})
print(child['a'])
print(child.parents['a'])
"#;
    assert_eq!(run_python(src), vec!["9", "1"]);
}
