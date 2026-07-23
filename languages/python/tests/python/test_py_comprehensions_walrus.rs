use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: walrus operator + comprehensions — :=, nested comprehensions, conditional expressions, unpacking targets
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_walrus_in_while_loop() {
    let src = r#"
import io

buf = io.StringIO("line1\nline2\nline3\n")
lines = []
while (line := buf.readline()):
    lines.append(line.strip())
print(lines)
"#;
    assert_eq!(run_python(src), vec!["['line1', 'line2', 'line3']"]);
}

#[test]
fn test_py_walrus_in_list_comprehension() {
    let src = r#"
data = [1, -2, 3, -4, 5]
results = [y for x in data if (y := x ** 2) > 4]
print(results)
"#;
    assert_eq!(run_python(src), vec!["[4, 9, 16, 25]"]);
}

#[test]
fn test_py_walrus_in_if_expression() {
    let src = r#"
import re

text = "The user is: alice@example.com"
if m := re.search(r"[\w.]+@[\w.]+\.\w+", text):
    print(f"Found email: {m.group()}")
else:
    print("No email found")
"#;
    assert_eq!(run_python(src), vec!["Found email: alice@example.com"]);
}

#[test]
fn test_py_walrus_avoid_recomputation() {
    let src = r#"
import math

data = [4, 9, 16, -1, 25]
# compute once in condition, reuse in body
results = []
for n in data:
    if (sq := math.sqrt(n)) > 3 if n >= 0 else False:
        results.append(sq)
print(results)
"#;
    assert_eq!(run_python(src), vec!["[3.0, 4.0, 5.0]"]);
}

#[test]
fn test_py_nested_list_comprehension_flatten() {
    let src = r#"
matrix = [[1, 2, 3], [4, 5, 6], [7, 8, 9]]
flat = [x for row in matrix for x in row]
print(flat)

# Only even values
evens = [x for row in matrix for x in row if x % 2 == 0]
print(evens)
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, 3, 4, 5, 6, 7, 8, 9]", "[2, 4, 6, 8]"]
    );
}

#[test]
fn test_py_dict_comprehension_inversion_and_grouping() {
    let src = r#"
words = ["apple", "ant", "bear", "bat", "cat"]
by_first = {
    letter: [w for w in words if w.startswith(letter)]
    for letter in set(w[0] for w in words)
}
print(sorted(by_first.keys()))
print(sorted(by_first["a"]))
print(sorted(by_first["b"]))
"#;
    assert_eq!(
        run_python(src),
        vec!["['a', 'b', 'c']", "['ant', 'apple']", "['bat', 'bear']"]
    );
}

#[test]
fn test_py_generator_expression_lazy_evaluation() {
    let src = r#"
import sys

numbers = range(10 ** 6)
gen_expr = (x ** 2 for x in numbers if x % 2 == 0)
list_comp = [x ** 2 for x in range(10) if x % 2 == 0]

gen_size = sys.getsizeof(gen_expr)
list_size = sys.getsizeof(list_comp)
print(gen_size < list_size)  # generator takes less memory
print(next(gen_expr))
print(list_comp)
"#;
    assert_eq!(run_python(src), vec!["True", "0", "[0, 4, 16, 36, 64]"]);
}

#[test]
fn test_py_conditional_expression_ternary() {
    let src = r#"
x = 10
label = "positive" if x > 0 else "non-positive"
print(label)

values = [abs(x) if x >= 0 else -x for x in range(-3, 4)]
print(values)

nested = "big" if x > 100 else "medium" if x > 10 else "small"
print(nested)
"#;
    assert_eq!(
        run_python(src),
        vec!["positive", "[3, 2, 1, 0, 1, 2, 3]", "small"]
    );
}

#[test]
fn test_py_unpacking_in_comprehension_target() {
    let src = r#"
pairs = [("alice", 30), ("bob", 25), ("carol", 35)]
names = [name for name, age in pairs if age > 25]
print(names)

age_map = {name: age for name, age in pairs}
print(age_map)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['alice', 'carol']",
            "{'alice': 30, 'bob': 25, 'carol': 35}"
        ]
    );
}

#[test]
fn test_py_walrus_set_comprehension() {
    let src = r#"
words = ["hello", "world", "hi", "hey", "world"]
# Use walrus to capture length while deduplicating
long_words = {(word, length) for word in words if (length := len(word)) > 3}
print(sorted(long_words))
"#;
    assert_eq!(run_python(src), vec!["[('hello', 5), ('world', 5)]"]);
}

#[test]
fn test_py_comprehension_with_walrus_early_exit() {
    let src = r#"
import re

emails = ["alice@example.com", "not-an-email", "bob@test.org", "broken@", "carol@domain.net"]
valid = [
    m.group()
    for email in emails
    if (m := re.fullmatch(r"[\w.]+@[\w]+\.\w+", email))
]
print(valid)
"#;
    assert_eq!(
        run_python(src),
        vec!["['alice@example.com', 'bob@test.org', 'carol@domain.net']"]
    );
}
