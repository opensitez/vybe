use super::helpers::run_python;

#[test]
fn test_python_itertools_permutations() {
    let src = r#"
import itertools
print(list(itertools.permutations([1, 2, 3], 2)))
"#;
    assert_eq!(
        run_python(src),
        vec!["[(1, 2), (1, 3), (2, 1), (2, 3), (3, 1), (3, 2)]"]
    );
}

#[test]
fn test_python_itertools_combinations() {
    let src = r#"
import itertools
print(list(itertools.combinations([1, 2, 3, 4], 2)))
"#;
    assert_eq!(
        run_python(src),
        vec!["[(1, 2), (1, 3), (1, 4), (2, 3), (2, 4), (3, 4)]"]
    );
}
