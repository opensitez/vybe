use super::helpers::run_python;

#[test]
fn test_python_random_seed_and_choice() {
    let src = r#"
import random
random.seed(123)
print(random.randint(1, 10))
print(random.choice(['a', 'b', 'c']))
"#;
    assert_eq!(run_python(src), vec!["1", "b"]);
}

#[test]
fn test_python_random_triangular_and_sample() {
    let src = r#"
import random
random.seed(1)
print(round(random.triangular(1, 10, 2), 1))
print(random.sample([1, 2, 3, 4], 2))
"#;
    assert_eq!(run_python(src), vec!["3.0", "[1, 4]"]);
}
