// Python sorted() — key, reverse, stability, custom objects
use super::helpers::run_python;

#[test]
fn test_sorted_basic() {
    let script = r#"
print(sorted([3, 1, 4, 1, 5, 9, 2, 6]))
"#;
    assert_eq!(run_python(script), vec!["[1, 1, 2, 3, 4, 5, 6, 9]"]);
}

#[test]
fn test_sorted_reverse() {
    let script = r#"
print(sorted([3, 1, 4, 1, 5], reverse=True))
"#;
    assert_eq!(run_python(script), vec!["[5, 4, 3, 1, 1]"]);
}

#[test]
fn test_sorted_key_function() {
    let script = r#"
words = ["banana", "Apple", "cherry", "date"]
print(sorted(words, key=str.lower))
"#;
    assert_eq!(run_python(script), vec!["['Apple', 'banana', 'cherry', 'date']"]);
}

#[test]
fn test_sorted_by_length_then_alpha() {
    let script = r#"
words = ["cat", "elephant", "bee", "ant", "dog"]
result = sorted(words, key=lambda w: (len(w), w))
print(result)
"#;
    assert_eq!(run_python(script), vec!["['ant', 'bee', 'cat', 'dog', 'elephant']"]);
}

#[test]
fn test_sorted_tuple_key() {
    let script = r#"
data = [(1, 'z'), (2, 'a'), (1, 'a'), (2, 'z')]
print(sorted(data))
"#;
    assert_eq!(run_python(script), vec!["[(1, 'a'), (1, 'z'), (2, 'a'), (2, 'z')]"]);
}

#[test]
fn test_sorted_does_not_mutate() {
    let script = r#"
original = [3, 1, 2]
result = sorted(original)
print(original)
print(result)
"#;
    assert_eq!(run_python(script), vec!["[3, 1, 2]", "[1, 2, 3]"]);
}

#[test]
fn test_sorted_custom_objects() {
    let script = r#"
class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age
    def __repr__(self):
        return self.name

people = [Person("Bob", 30), Person("Alice", 25), Person("Carol", 35)]
by_age = sorted(people, key=lambda p: p.age)
print([p.name for p in by_age])
"#;
    assert_eq!(run_python(script), vec!["['Alice', 'Bob', 'Carol']"]);
}

#[test]
fn test_sorted_strings_lexicographic() {
    let script = r#"
print(sorted(["banana", "apple", "cherry"]))
"#;
    assert_eq!(run_python(script), vec!["['apple', 'banana', 'cherry']"]);
}
