// Python format mini-language — str.format(), format(), format_map()
use super::helpers::run_python;

#[test]
fn test_format_positional() {
    let script = r#"
print("{} {} {}".format(1, 2, 3))
print("{0} and {1}".format("hello", "world"))
"#;
    assert_eq!(run_python(script), vec!["1 2 3", "hello and world"]);
}

#[test]
fn test_format_keyword() {
    let script = r#"
print("{name} is {age}".format(name="Alice", age=30))
"#;
    assert_eq!(run_python(script), vec!["Alice is 30"]);
}

#[test]
fn test_format_index_reuse() {
    let script = r#"
print("{0} {0} {1}".format("ha", "!"))
"#;
    assert_eq!(run_python(script), vec!["ha ha !"]);
}

#[test]
fn test_format_float_spec() {
    let script = r#"
print("{:.2f}".format(3.14159))
print("{:8.3f}".format(1.5))
"#;
    assert_eq!(run_python(script), vec!["3.14", "   1.500"]);
}

#[test]
fn test_format_integer_radix() {
    let script = r#"
print("{:b}".format(10))
print("{:x}".format(255))
print("{:#o}".format(8))
"#;
    assert_eq!(run_python(script), vec!["1010", "ff", "0o10"]);
}

#[test]
fn test_format_map() {
    let script = r#"
data = {"name": "Bob", "score": 95}
print("{name}: {score}".format_map(data))
"#;
    assert_eq!(run_python(script), vec!["Bob: 95"]);
}

#[test]
fn test_builtin_format_function() {
    let script = r#"
print(format(3.14159, ".2f"))
print(format(255, "08b"))
"#;
    assert_eq!(run_python(script), vec!["3.14", "11111111"]);
}

#[test]
fn test_format_dict_access() {
    let script = r#"
person = {"name": "Carol"}
print("{p[name]}".format(p=person))
"#;
    assert_eq!(run_python(script), vec!["Carol"]);
}

#[test]
fn test_format_percent_literal() {
    let script = r#"
print("{:.1%}".format(0.875))
"#;
    assert_eq!(run_python(script), vec!["87.5%"]);
}
