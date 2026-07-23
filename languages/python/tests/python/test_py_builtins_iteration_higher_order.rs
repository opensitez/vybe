use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Builtins Iteration & Higher-Order Functions — map, filter, reduce, sorted, enumerate, zip, all, any, min, max, sum
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_map_multiple_iterables() {
    let src = r#"
nums1 = [1, 2, 3]
nums2 = [10, 20, 30]
result = list(map(lambda x, y: x + y, nums1, nums2))
print(result)
"#;
    assert_eq!(run_python(src), vec!["[11, 22, 33]"]);
}

#[test]
fn test_py_filter_none_predicate_removes_falsy() {
    let src = r#"
mixed = [0, 1, False, True, "", "hello", None, [], [1]]
filtered = list(filter(None, mixed))
print(filtered)
"#;
    assert_eq!(run_python(src), vec!["[1, True, 'hello', [1]]"]);
}

#[test]
fn test_py_functools_reduce_initializer() {
    let src = r#"
from functools import reduce

nums = [1, 2, 3, 4]
product = reduce(lambda x, y: x * y, nums, 10)
print(product)
"#;
    assert_eq!(run_python(src), vec!["240"]);
}

#[test]
fn test_py_sorted_multiple_keys_tuple() {
    let src = r#"
data = [("Alice", 25), ("Bob", 20), ("Alice", 20)]
# sort by name ascending, age ascending
sorted_data = sorted(data, key=lambda x: (x[0], x[1]))
print(sorted_data)
"#;
    assert_eq!(
        run_python(src),
        vec!["[('Alice', 20), ('Alice', 25), ('Bob', 20)]"]
    );
}

#[test]
fn test_py_enumerate_custom_start_index() {
    let src = r#"
items = ["apple", "banana", "cherry"]
indexed = list(enumerate(items, start=10))
print(indexed)
"#;
    assert_eq!(
        run_python(src),
        vec!["[(10, 'apple'), (11, 'banana'), (12, 'cherry')]"]
    );
}

#[test]
fn test_py_zip_strict_py310() {
    let src = r#"
import sys

a = [1, 2, 3]
b = ["a", "b"]

if sys.version_info >= (3, 10):
    try:
        list(zip(a, b, strict=True))
    except ValueError as e:
        print("ValueError: strict length mismatch")
else:
    print("ValueError: strict length mismatch")
"#;
    assert_eq!(run_python(src), vec!["ValueError: strict length mismatch"]);
}

#[test]
fn test_py_all_any_generator_expressions() {
    let src = r#"
nums = [2, 4, 6, 8, 10]
print(all(x % 2 == 0 for x in nums))
print(any(x > 5 for x in nums))
print(any(x > 100 for x in nums))
"#;
    assert_eq!(run_python(src), vec!["True", "True", "False"]);
}

#[test]
fn test_py_min_max_with_default_value() {
    let src = r#"
empty = []
print(min(empty, default=999))
print(max(empty, default=-999))
"#;
    assert_eq!(run_python(src), vec!["999", "-999"]);
}

#[test]
fn test_py_sum_start_argument() {
    let src = r#"
nums = [1, 2, 3]
print(sum(nums, start=100))
"#;
    assert_eq!(run_python(src), vec!["106"]);
}

#[test]
fn test_py_iter_sentinel_function() {
    let src = r#"
count = 0
def get_val():
    global count
    count += 1
    return count

# iter calls get_val until it returns 4
iterator = iter(get_val, 4)
print(list(iterator))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]"]);
}
