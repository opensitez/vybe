use super::helpers::run_python;

#[test]
fn test_python_itertools_chain_repeat() {
    let src = r#"
import itertools
print(list(itertools.chain([1, 2], [3, 4])))
print(list(itertools.repeat('x', 3)))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4]", "['x', 'x', 'x']"]);
}

#[test]
fn test_python_itertools_accumulate() {
    let src = r#"
import itertools
print(list(itertools.accumulate([1, 2, 3, 4], initial=0)))
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 3, 6, 10]"]);
}
