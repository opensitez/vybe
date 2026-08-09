use super::helpers::run_python;

#[test]
fn test_operator_itemgetter_single_index() {
    let out = run_python(
        r##"
import operator
get_second = operator.itemgetter(1)
data = ["alpha", "beta", "gamma"]
print(get_second(data))
"##,
    );
    assert_eq!(out, vec!["beta"]);
}

#[test]
fn test_operator_itemgetter_multiple_indices() {
    let out = run_python(
        r##"
import operator
get_first_third = operator.itemgetter(0, 2)
data = (10, 20, 30, 40)
print(get_first_third(data))
"##,
    );
    assert_eq!(out, vec!["(10, 30)"]);
}

#[test]
fn test_operator_itemgetter_dict_keys() {
    let out = run_python(
        r##"
import operator
get_ab = operator.itemgetter("a", "b")
d = {"a": 100, "b": 200, "c": 300}
print(get_ab(d))
"##,
    );
    assert_eq!(out, vec!["(100, 200)"]);
}

#[test]
fn test_operator_attrgetter_single_attr() {
    let out = run_python(
        r##"
import operator
class Person:
    def __init__(self, name):
        self.name = name

get_name = operator.attrgetter("name")
p = Person("Alice")
print(get_name(p))
"##,
    );
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn test_operator_attrgetter_nested_attr() {
    let out = run_python(
        r##"
import operator
class Node:
    def __init__(self, child=None, val=0):
        self.child = child
        self.val = val

get_grandchild_val = operator.attrgetter("child.child.val")
tree = Node(child=Node(child=Node(val=42)))
print(get_grandchild_val(tree))
"##,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn test_operator_methodcaller_no_args() {
    let out = run_python(
        r##"
import operator
to_upper = operator.methodcaller("upper")
text = "hello world"
print(to_upper(text))
"##,
    );
    assert_eq!(out, vec!["HELLO WORLD"]);
}

#[test]
fn test_operator_methodcaller_with_args() {
    let out = run_python(
        r##"
import operator
repl = operator.methodcaller("replace", "foo", "bar")
s = "foo_test_foo"
print(repl(s))
"##,
    );
    assert_eq!(out, vec!["bar_test_bar"]);
}

#[test]
fn test_operator_arithmetic_add_sub_mul() {
    let out = run_python(
        r##"
import operator
print(operator.add(15, 27))
print(operator.sub(50, 18))
print(operator.mul(6, 7))
"##,
    );
    assert_eq!(out, vec!["42", "32", "42"]);
}

#[test]
fn test_operator_truth_and_not_() {
    let out = run_python(
        r##"
import operator
print(operator.truth([1]))
print(operator.truth([]))
print(operator.not_(True))
print(operator.not_(0))
"##,
    );
    assert_eq!(out, vec!["True", "False", "False", "True"]);
}

#[test]
fn test_operator_contains() {
    let out = run_python(
        r##"
import operator
items = ["apple", "banana", "cherry"]
print(operator.contains(items, "banana"))
print(operator.contains(items, "orange"))
"##,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_operator_count_of() {
    let out = run_python(
        r##"
import operator
nums = [1, 2, 2, 3, 2, 4]
print(operator.countOf(nums, 2))
print(operator.countOf(nums, 5))
"##,
    );
    assert_eq!(out, vec!["3", "0"]);
}

#[test]
fn test_operator_index_of() {
    let out = run_python(
        r##"
import operator
letters = ["a", "b", "c", "d"]
print(operator.indexOf(letters, "c"))
try:
    operator.indexOf(letters, "z")
except ValueError:
    print("NOT_FOUND")
"##,
    );
    assert_eq!(out, vec!["2", "NOT_FOUND"]);
}

#[test]
fn test_operator_concat() {
    let out = run_python(
        r##"
import operator
print(operator.concat("hello ", "world"))
print(operator.concat([1, 2], [3, 4]))
"##,
    );
    assert_eq!(out, vec!["hello world", "[1, 2, 3, 4]"]);
}

#[test]
fn test_operator_length_hint() {
    let out = run_python(
        r##"
import operator
print(operator.length_hint([10, 20, 30]))
print(operator.length_hint(iter([1, 2]), default=99))
"##,
    );
    assert_eq!(out, vec!["3", "2"]);
}

#[test]
fn test_operator_setitem_delitem() {
    let out = run_python(
        r##"
import operator
d = {"x": 1}
operator.setitem(d, "y", 2)
print(d)
operator.delitem(d, "x")
print(d)
"##,
    );
    assert_eq!(out, vec!["{'x': 1, 'y': 2}", "{'y': 2}"]);
}

#[test]
fn test_operator_iadd_isub() {
    let out = run_python(
        r##"
import operator
lst = [1, 2]
lst = operator.iadd(lst, [3, 4])
print(lst)
val = 10
val = operator.isub(val, 3)
print(val)
"##,
    );
    assert_eq!(out, vec!["[1, 2, 3, 4]", "7"]);
}

#[test]
fn test_operator_abs_neg_pos() {
    let out = run_python(
        r##"
import operator
print(operator.abs(-15))
print(operator.neg(7))
print(operator.pos(-5))
"##,
    );
    assert_eq!(out, vec!["15", "-7", "-5"]);
}

#[test]
fn test_operator_index() {
    let out = run_python(
        r##"
import operator
class CustomIndex:
    def __index__(self):
        return 7

ci = CustomIndex()
print(operator.index(ci))
"##,
    );
    assert_eq!(out, vec!["7"]);
}

#[test]
fn test_operator_eq_ne_lt_gt() {
    let out = run_python(
        r##"
import operator
print(operator.eq(5, 5))
print(operator.ne(5, 3))
print(operator.lt(2, 8))
print(operator.gt(10, 4))
"##,
    );
    assert_eq!(out, vec!["True", "True", "True", "True"]);
}

#[test]
fn test_operator_matmul() {
    let out = run_python(
        r##"
import operator
class Matrix:
    def __init__(self, val):
        self.val = val
    def __matmul__(self, other):
        return Matrix(self.val * other.val)
    def __repr__(self):
        return f"M({self.val})"

m1 = Matrix(3)
m2 = Matrix(4)
print(operator.matmul(m1, m2))
"##,
    );
    assert_eq!(out, vec!["M(12)"]);
}
