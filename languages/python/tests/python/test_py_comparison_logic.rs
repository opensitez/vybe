use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Comparison & Logic — chained comparisons, is vs ==, truthiness, short-circuiting, bitwise vs logical
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_chained_comparisons() {
    let src = r#"
x = 5
print(1 < x < 10)
print(1 < x > 10)
print(0 <= x <= 5 == 5)
print(10 > x > 0 < 2)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True", "True"]);
}

#[test]
fn test_py_identity_is_vs_equality_eq() {
    let src = r#"
a = [1, 2, 3]
b = [1, 2, 3]
c = a

print(a == b)
print(a is b)
print(a is c)
print(a is not b)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True", "True"]);
}

#[test]
fn test_py_singleton_identity_none_true_false() {
    let src = r#"
val = None
print(val is None)
print(val is not None)

flag = True
print(flag is True)
print(flag is False)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True", "False"]);
}

#[test]
fn test_py_short_circuit_and_or_evaluation() {
    let src = r#"
log = []

def side_effect(val, label):
    log.append(label)
    return val

res1 = side_effect(False, "A") and side_effect(True, "B")
print(res1)
print(log)

log.clear()
res2 = side_effect(True, "X") or side_effect(False, "Y")
print(res2)
print(log)
"#;
    assert_eq!(run_python(src), vec!["False", "['A']", "True", "['X']"]);
}

#[test]
fn test_py_and_or_operator_returned_values() {
    let src = r#"
# and / or return the actual operand, not bool
print([] or "default")
print("first" or "second")
print("first" and "second")
print("" and "second")
"#;
    assert_eq!(run_python(src), vec!["default", "first", "second", ""]);
}

#[test]
fn test_py_custom_truthiness_bool_len_dunders() {
    let src = r#"
class CustomBool:
    def __init__(self, val):
        self.val = val

    def __bool__(self):
        return self.val > 0

class CustomLen:
    def __init__(self, items):
        self.items = items

    def __len__(self):
        return len(self.items)

print(bool(CustomBool(5)))
print(bool(CustomBool(-1)))

print(bool(CustomLen([1, 2])))
print(bool(CustomLen([])))
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True", "False"]);
}

#[test]
fn test_py_not_operator_boolean_inversion() {
    let src = r#"
print(not True)
print(not False)
print(not 0)
print(not 10)
print(not "")
print(not "text")
"#;
    assert_eq!(
        run_python(src),
        vec!["False", "True", "True", "False", "True", "False"]
    );
}

#[test]
fn test_py_in_and_not_in_membership_operators() {
    let src = r#"
lst = [1, 2, 3]
d = {"a": 10}
s = "hello"

print(2 in lst)
print(5 not in lst)
print("a" in d)
print(10 not in d)  # checks keys
print("ell" in s)
"#;
    assert_eq!(
        run_python(src),
        vec!["True", "True", "True", "True", "True"]
    );
}

#[test]
fn test_py_comparison_precedence_with_logical_operators() {
    let src = r#"
x = 10
y = 20
z = 30

# Comparison operators have higher precedence than logical and/or/not
print(x < y and y < z)
print(not x == y)
print(x == y or z > y)
"#;
    assert_eq!(run_python(src), vec!["True", "True", "True"]);
}

#[test]
fn test_py_int_string_interop_comparison() {
    let src = r#"
print(10 == "10")
print(10 != "10")
try:
    10 < "10"
except TypeError:
    print("TypeError: unorderable types")
"#;
    assert_eq!(
        run_python(src),
        vec!["False", "True", "TypeError: unorderable types"]
    );
}
