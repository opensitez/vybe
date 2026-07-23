use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: operator module — operator functions, attrgetter, itemgetter, methodcaller, comparison operators as functions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_operator_arithmetic() {
    let src = r#"
import operator

print(operator.add(10, 5))
print(operator.sub(10, 5))
print(operator.mul(10, 5))
print(operator.truediv(10, 4))
print(operator.floordiv(10, 3))
print(operator.mod(10, 3))
print(operator.pow(2, 8))
"#;
    assert_eq!(
        run_python(src),
        vec!["15", "5", "50", "2.5", "3", "1", "256"]
    );
}

#[test]
fn test_py_operator_comparison() {
    let src = r#"
import operator

print(operator.lt(3, 5))
print(operator.le(5, 5))
print(operator.gt(7, 3))
print(operator.ge(3, 3))
print(operator.eq(4, 4))
print(operator.ne(4, 5))
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_py_operator_logical() {
    let src = r#"
import operator

print(operator.and_(True, False))
print(operator.or_(True, False))
print(operator.not_(False))
print(operator.truth(0))
print(operator.truth(1))
"#;
    assert_eq!(
        run_python(src),
        vec!["False", "True", "True", "False", "True"]
    );
}

#[test]
fn test_py_operator_bitwise() {
    let src = r#"
import operator

print(operator.and_(0b1100, 0b1010))
print(operator.or_(0b1100, 0b1010))
print(operator.xor(0b1100, 0b1010))
print(operator.invert(~5 - 1))
print(operator.lshift(1, 4))
print(operator.rshift(256, 4))
"#;
    assert_eq!(run_python(src), vec!["8", "14", "6", "4", "16", "16"]);
}

#[test]
fn test_py_operator_itemgetter() {
    let src = r#"
from operator import itemgetter

data = [{"name": "Bob", "age": 25}, {"name": "Alice", "age": 30}, {"name": "Charlie", "age": 20}]
get_name = itemgetter("name")
print(get_name(data[0]))

sorted_by_age = sorted(data, key=itemgetter("age"))
print([d["name"] for d in sorted_by_age])

# Multiple keys
get_both = itemgetter("name", "age")
print(get_both(data[0]))
"#;
    assert_eq!(
        run_python(src),
        vec!["Bob", "['Charlie', 'Bob', 'Alice']", "('Bob', 25)"]
    );
}

#[test]
fn test_py_operator_attrgetter() {
    let src = r#"
from operator import attrgetter

class Person:
    def __init__(self, name, age):
        self.name = name
        self.age = age

people = [Person("Bob", 25), Person("Alice", 30), Person("Charlie", 20)]
get_name = attrgetter("name")
print(get_name(people[0]))

sorted_by_age = sorted(people, key=attrgetter("age"))
print([p.name for p in sorted_by_age])
"#;
    assert_eq!(run_python(src), vec!["Bob", "['Charlie', 'Bob', 'Alice']"]);
}

#[test]
fn test_py_operator_methodcaller() {
    let src = r#"
from operator import methodcaller

upper = methodcaller("upper")
words = ["hello", "world", "python"]
print(list(map(upper, words)))

strip_and_split = methodcaller("split", ",")
csv_row = "a,b,c"
print(strip_and_split(csv_row))
"#;
    assert_eq!(
        run_python(src),
        vec!["['HELLO', 'WORLD', 'PYTHON']", "['a', 'b', 'c']"]
    );
}

#[test]
fn test_py_operator_getitem_setitem_delitem() {
    let src = r#"
import operator

d = {"a": 1, "b": 2}
print(operator.getitem(d, "a"))
operator.setitem(d, "c", 3)
print(d)
operator.delitem(d, "a")
print(d)

lst = [10, 20, 30]
print(operator.getitem(lst, 1))
"#;
    assert_eq!(
        run_python(src),
        vec!["1", "{'a': 1, 'b': 2, 'c': 3}", "{'b': 2, 'c': 3}", "20"]
    );
}

#[test]
fn test_py_operator_concat_contains() {
    let src = r#"
import operator

print(operator.concat([1, 2], [3, 4]))
print(operator.concat("Hello", " World"))
print(operator.contains([1, 2, 3], 2))
print(operator.contains([1, 2, 3], 5))
print(operator.length_hint([1, 2, 3]))
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, 3, 4]", "Hello World", "True", "False", "3"]
    );
}

#[test]
fn test_py_operator_in_functional_pipeline() {
    let src = r#"
import operator
from functools import reduce

numbers = [1, 2, 3, 4, 5]
total = reduce(operator.add, numbers)
product = reduce(operator.mul, numbers)
print(total)
print(product)

# Build matrix of comparisons
pairs = [(1, 2), (5, 3), (4, 4)]
results = [operator.lt(a, b) for a, b in pairs]
print(results)
"#;
    assert_eq!(run_python(src), vec!["15", "120", "[True, False, False]"]);
}
