// Python extended unpacking — starred, nested, swap, for loops
use super::helpers::run_python;

#[test]
fn test_starred_first() {
    let script = r#"
first, *rest = [1, 2, 3, 4, 5]
print(first)
print(rest)
"#;
    assert_eq!(run_python(script), vec!["1", "[2, 3, 4, 5]"]);
}

#[test]
fn test_starred_last() {
    let script = r#"
*head, last = [1, 2, 3, 4, 5]
print(head)
print(last)
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3, 4]", "5"]);
}

#[test]
fn test_starred_middle() {
    let script = r#"
first, *middle, last = range(6)
print(first)
print(middle)
print(last)
"#;
    assert_eq!(run_python(script), vec!["0", "[1, 2, 3, 4]", "5"]);
}

#[test]
fn test_nested_tuple_unpack() {
    let script = r#"
(a, b), c = (1, 2), 3
print(a, b, c)
"#;
    assert_eq!(run_python(script), vec!["1 2 3"]);
}

#[test]
fn test_swap_in_place() {
    let script = r#"
x, y = 10, 20
x, y = y, x
print(x, y)
"#;
    assert_eq!(run_python(script), vec!["20 10"]);
}

#[test]
fn test_unpack_in_for_loop() {
    let script = r#"
pairs = [(1, 'a'), (2, 'b'), (3, 'c')]
for num, letter in pairs:
    print(num, letter)
"#;
    assert_eq!(run_python(script), vec!["1 a", "2 b", "3 c"]);
}

#[test]
fn test_starred_in_function_call() {
    let script = r#"
def add(a, b, c):
    return a + b + c

args = [1, 2, 3]
print(add(*args))
"#;
    assert_eq!(run_python(script), vec!["6"]);
}

#[test]
fn test_dict_unpack_merge() {
    let script = r#"
d1 = {"a": 1, "b": 2}
d2 = {"c": 3, "d": 4}
merged = {**d1, **d2}
print(sorted(merged.items()))
"#;
    assert_eq!(run_python(script), vec!["[('a', 1), ('b', 2), ('c', 3), ('d', 4)]"]);
}

#[test]
fn test_star_in_list_literal() {
    let script = r#"
a = [1, 2]
b = [3, 4]
c = [*a, *b, 5]
print(c)
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3, 4, 5]"]);
}
