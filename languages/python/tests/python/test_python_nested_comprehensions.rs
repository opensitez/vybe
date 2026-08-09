// Python nested comprehensions — 2D, filtering, flattening, dict/set from nested
use super::helpers::run_python;

#[test]
fn test_nested_list_comprehension_matrix() {
    let script = r#"
matrix = [[i * j for j in range(1, 4)] for i in range(1, 4)]
print(matrix)
"#;
    assert_eq!(
        run_python(script),
        vec!["[[1, 2, 3], [2, 4, 6], [3, 6, 9]]"]
    );
}

#[test]
fn test_nested_flatten() {
    let script = r#"
nested = [[1, 2, 3], [4, 5], [6, 7, 8, 9]]
flat = [x for sublist in nested for x in sublist]
print(flat)
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3, 4, 5, 6, 7, 8, 9]"]);
}

#[test]
fn test_nested_with_filter() {
    let script = r#"
pairs = [(x, y) for x in range(4) for y in range(4) if x != y and x + y == 3]
print(sorted(pairs))
"#;
    assert_eq!(run_python(script), vec!["[(0, 3), (1, 2), (2, 1), (3, 0)]"]);
}

#[test]
fn test_nested_dict_comprehension() {
    let script = r#"
keys = ["a", "b", "c"]
vals = [1, 2, 3]
d = {k: v for k, v in zip(keys, vals)}
print(d)
"#;
    assert_eq!(run_python(script), vec!["{'a': 1, 'b': 2, 'c': 3}"]);
}

#[test]
fn test_nested_set_comprehension() {
    let script = r#"
nums = [1, 2, 3, 1, 2, 4]
unique_squares = {x ** 2 for x in nums}
print(sorted(unique_squares))
"#;
    assert_eq!(run_python(script), vec!["[1, 4, 9, 16]"]);
}

#[test]
fn test_nested_generator_expression() {
    let script = r#"
total = sum(x * y for x in range(1, 4) for y in range(1, 4) if x == y)
print(total)  # 1*1 + 2*2 + 3*3 = 14
"#;
    assert_eq!(run_python(script), vec!["14"]);
}

#[test]
fn test_nested_conditional_comprehension() {
    let script = r#"
result = ["even" if x % 2 == 0 else "odd" for x in range(6)]
print(result)
"#;
    assert_eq!(
        run_python(script),
        vec!["['even', 'odd', 'even', 'odd', 'even', 'odd']"]
    );
}

#[test]
fn test_nested_comprehension_with_walrus() {
    let script = r#"
data = [1, -2, 3, -4, 5]
pos = [y for x in data if (y := abs(x)) > 2]
print(pos)
"#;
    assert_eq!(run_python(script), vec!["[3, 4, 5]"]);
}
