use super::helpers::run_python;

#[test]
fn test_python_itertools_groupby_keyed() {
    let src = r#"
import itertools
items = [('a', 1), ('a', 2), ('b', 1), ('b', 3)]
for key, group in itertools.groupby(items, lambda x: x[0]):
    vals = [v[1] for v in group]
    print(f"{key}:{sum(vals)}")
"#;
    assert_eq!(run_python(src), vec!["a:3", "b:4"]);
}

#[test]
fn test_python_itertools_pairwise_fallback() {
    let src = r#"
import itertools
pairs = list(itertools.pairwise('abcd'))
print(pairs)
"#;
    assert_eq!(
        run_python(src),
        vec!["[('a', 'b'), ('b', 'c'), ('c', 'd')]"]
    );
}
