use super::helpers::run_python;

#[test]
fn test_python_defaultdict_int_factory() {
    let src = r#"
from collections import defaultdict
count = defaultdict(int)
count['a'] += 2
count['b'] += 1
print(count['a'])
print(count['c'])
print(count['b'])
"#;
    assert_eq!(run_python(src), vec!["2", "0", "1"]);
}

#[test]
fn test_python_defaultdict_factory_list() {
    let src = r#"
from collections import defaultdict
bags = defaultdict(list)
bags['x'].append(1)
bags['x'].append(2)
print(bags['x'])
print(sorted(bags.items()))
"#;
    assert_eq!(run_python(src), vec!["[1, 2]", "[('x', [1, 2])]"]);
}
