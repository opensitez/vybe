// Python operator module — functional versions of operators, itemgetter, attrgetter
use super::helpers::run_python;

#[test]
fn test_operator_add_mul() {
    let script = r#"
import operator
print(operator.add(3, 4))
print(operator.mul(5, 6))
print(operator.sub(10, 3))
print(operator.truediv(7, 2))
"#;
    assert_eq!(run_python(script), vec!["7", "30", "7", "3.5"]);
}

#[test]
fn test_operator_comparison() {
    let script = r#"
import operator
print(operator.lt(1, 2))
print(operator.ge(5, 5))
print(operator.eq(3, 3))
print(operator.ne(1, 2))
"#;
    assert_eq!(run_python(script), vec!["True", "True", "True", "True"]);
}

#[test]
fn test_operator_itemgetter() {
    let script = r#"
import operator
data = [{'name': 'Bob', 'age': 30}, {'name': 'Alice', 'age': 25}]
by_age = sorted(data, key=operator.itemgetter('age'))
print([d['name'] for d in by_age])
"#;
    assert_eq!(run_python(script), vec!["['Alice', 'Bob']"]);
}

#[test]
fn test_operator_attrgetter() {
    let script = r#"
import operator

class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

people = [Person('Bob', 30), Person('Alice', 25), Person('Carol', 35)]
by_name = sorted(people, key=operator.attrgetter('name'))
print([p.name for p in by_name])
"#;
    assert_eq!(run_python(script), vec!["['Alice', 'Bob', 'Carol']"]);
}

#[test]
fn test_operator_methodcaller() {
    let script = r#"
import operator
words = ["hello", "WORLD", "Python"]
upper = list(map(operator.methodcaller('upper'), words))
print(upper)
"#;
    assert_eq!(run_python(script), vec!["['HELLO', 'WORLD', 'PYTHON']"]);
}

#[test]
fn test_operator_getitem_setitem() {
    let script = r#"
import operator
lst = [10, 20, 30]
print(operator.getitem(lst, 1))
operator.setitem(lst, 2, 99)
print(lst)
"#;
    assert_eq!(run_python(script), vec!["20", "[10, 20, 99]"]);
}

#[test]
fn test_operator_logical() {
    let script = r#"
import operator
print(operator.and_(0b1010, 0b1100))
print(operator.or_(0b1010, 0b1100))
print(operator.xor(0b1010, 0b1100))
print(operator.not_(False))
"#;
    assert_eq!(run_python(script), vec!["8", "14", "6", "True"]);
}

#[test]
fn test_operator_concat_length() {
    let script = r#"
import operator
print(operator.concat([1, 2], [3, 4]))
print(operator.length_hint([1, 2, 3]))
"#;
    assert_eq!(run_python(script), vec!["[1, 2, 3, 4]", "3"]);
}
