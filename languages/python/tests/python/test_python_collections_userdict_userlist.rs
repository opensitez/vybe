use super::helpers::run_python;

#[test]
fn test_python_userdict_default_behaviour() {
    let src = r#"
from collections import UserDict
m = UserDict({'x': 1})
m['y'] = 2
print(m['x'])
print(m['y'])
"#;
    assert_eq!(run_python(src), vec!["1", "2"]);
}

#[test]
fn test_python_userlist_basic_ops() {
    let src = r#"
from collections import UserList
ul = UserList([1, 2])
ul.append(3)
ul.extend([4, 5])
print(ul.data)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4, 5]"]);
}
