// Python list advanced operations — sort stability, bisect, deque as list, list comprehension edge cases
use super::helpers::run_python;

#[test]
fn test_list_sort_stability() {
    let script = r#"
data = [(1, 'b'), (2, 'a'), (1, 'a'), (2, 'b')]
data.sort(key=lambda x: x[0])
print(data)
"#;
    assert_eq!(
        run_python(script),
        vec!["[(1, 'b'), (1, 'a'), (2, 'a'), (2, 'b')]"]
    );
}

#[test]
fn test_list_sort_reverse() {
    let script = r#"
nums = [3, 1, 4, 1, 5, 9, 2, 6]
nums.sort(reverse=True)
print(nums)
"#;
    assert_eq!(run_python(script), vec!["[9, 6, 5, 4, 3, 2, 1, 1]"]);
}

#[test]
fn test_list_copy_vs_slice() {
    let script = r#"
a = [1, [2, 3], 4]
b = a.copy()
b[0] = 99
b[1][0] = 99
print(a[0])       # shallow: outer unchanged
print(a[1][0])    # shallow: inner shared
"#;
    assert_eq!(run_python(script), vec!["1", "99"]);
}

#[test]
fn test_list_insert_remove_pop() {
    let script = r#"
lst = [1, 2, 3, 4, 5]
lst.insert(2, 99)
lst.remove(4)
popped = lst.pop(0)
print(popped)
print(lst)
"#;
    assert_eq!(run_python(script), vec!["1", "[2, 99, 3, 5]"]);
}

#[test]
fn test_list_extend_vs_append() {
    let script = r#"
a = [1, 2]
b = [1, 2]
a.append([3, 4])
b.extend([3, 4])
print(len(a))
print(len(b))
print(a[-1])
"#;
    assert_eq!(run_python(script), vec!["3", "4", "[3, 4]"]);
}

#[test]
fn test_list_count_and_index() {
    let script = r#"
lst = [1, 2, 3, 2, 1, 2]
print(lst.count(2))
print(lst.index(3))
print(lst.index(2, 2))  # start from index 2
"#;
    assert_eq!(run_python(script), vec!["3", "2", "3"]);
}

#[test]
fn test_list_multiplication() {
    let script = r#"
a = [0] * 5
print(a)
b = [[]] * 3
b[0].append(1)
print(b)  # all share same list reference
"#;
    assert_eq!(
        run_python(script),
        vec!["[0, 0, 0, 0, 0]", "[[1], [1], [1]]"]
    );
}

#[test]
fn test_list_clear_and_del() {
    let script = r#"
lst = [1, 2, 3, 4, 5]
del lst[1:3]
print(lst)
lst.clear()
print(lst)
"#;
    assert_eq!(run_python(script), vec!["[1, 4, 5]", "[]"]);
}

#[test]
fn test_list_reverse_in_place() {
    let script = r#"
lst = list(range(5))
lst.reverse()
print(lst)
"#;
    assert_eq!(run_python(script), vec!["[4, 3, 2, 1, 0]"]);
}
