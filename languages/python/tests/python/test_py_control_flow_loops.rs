use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Control Flow & Loops — for-else, while-else, break, continue, nested loops, conditional expressions
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_for_else_loop_search() {
    let src = r#"
def find_even(nums):
    for n in nums:
        if n % 2 == 0:
            print(f"Found even: {n}")
            break
    else:
        print("No even number found")

find_even([1, 3, 5, 6, 7])
find_even([1, 3, 5, 7])
"#;
    assert_eq!(
        run_python(src),
        vec!["Found even: 6", "No even number found"]
    );
}

#[test]
fn test_py_while_else_loop() {
    let src = r#"
count = 3
while count > 0:
    count -= 1
else:
    print("while-else reached")

count = 3
while count > 0:
    if count == 2:
        break
    count -= 1
else:
    print("while-else skipped on break")
"#;
    assert_eq!(run_python(src), vec!["while-else reached"]);
}

#[test]
fn test_py_nested_loop_break_with_flag() {
    let src = r#"
matrix = [[1, 2], [3, 4], [5, 6]]
found = None
for row in matrix:
    for val in row:
        if val == 4:
            found = val
            break
    if found is not None:
        break

print(found)
"#;
    assert_eq!(run_python(src), vec!["4"]);
}

#[test]
fn test_py_continue_in_loops() {
    let src = r#"
evens = []
for i in range(10):
    if i % 2 != 0:
        continue
    evens.append(i)

print(evens)
"#;
    assert_eq!(run_python(src), vec!["[0, 2, 4, 6, 8]"]);
}

#[test]
fn test_py_loop_unpacking_structures() {
    let src = r#"
pairs = [("a", 1), ("b", 2), ("c", 3)]
out = []
for k, v in pairs:
    out.append(f"{k}:{v}")

print(", ".join(out))
"#;
    assert_eq!(run_python(src), vec!["a:1, b:2, c:3"]);
}

#[test]
fn test_py_ternary_conditional_expression_chaining() {
    let src = r#"
def classify(score):
    return "A" if score >= 90 else ("B" if score >= 80 else ("C" if score >= 70 else "F"))

print(classify(95))
print(classify(85))
print(classify(75))
print(classify(50))
"#;
    assert_eq!(run_python(src), vec!["A", "B", "C", "F"]);
}

#[test]
fn test_py_pass_statement_no_op() {
    let src = r#"
class Empty:
    pass

def stub():
    pass

for _ in range(3):
    pass

print("pass works")
"#;
    assert_eq!(run_python(src), vec!["pass works"]);
}

#[test]
fn test_py_loop_mutation_during_iteration_safety() {
    let src = r#"
lst = [1, 2, 3, 4, 5]
# Iterating over a copy to safely mutate original
for item in lst[:]:
    if item % 2 == 0:
        lst.remove(item)

print(lst)
"#;
    assert_eq!(run_python(src), vec!["[1, 3, 5]"]);
}

#[test]
fn test_py_zip_parallel_iteration_loops() {
    let src = r#"
names = ["Alice", "Bob", "Charlie"]
ages = [25, 30, 35]

for name, age in zip(names, ages):
    print(f"{name} is {age}")
"#;
    assert_eq!(
        run_python(src),
        vec!["Alice is 25", "Bob is 30", "Charlie is 35"]
    );
}

#[test]
fn test_py_reversed_loop_iteration() {
    let src = r#"
items = ["first", "second", "third"]
out = []
for item in reversed(items):
    out.append(item)

print(out)
"#;
    assert_eq!(run_python(src), vec!["['third', 'second', 'first']"]);
}
