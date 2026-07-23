use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: List Sequence Operations — slicing assignment, in-place sort, list extension, repetition, shallow copy
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_list_slice_assignment_and_replacement() {
    let src = r#"
lst = [1, 2, 3, 4, 5]
lst[1:4] = [20, 30]
print(lst)

lst[::2] = [100, 200]
print(lst)
"#;
    assert_eq!(run_python(src), vec!["[1, 20, 30, 5]", "[100, 20, 200, 5]"]);
}

#[test]
fn test_py_list_slice_deletion() {
    let src = r#"
lst = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9]
del lst[1:8:2]
print(lst)
"#;
    assert_eq!(run_python(src), vec!["[0, 2, 4, 6, 8, 9]"]);
}

#[test]
fn test_py_list_in_place_sort_and_reverse() {
    let src = r#"
words = ["banana", "apple", "cherry", "date"]
words.sort(key=len)
print(words)

words.sort(key=len, reverse=True)
print(words)

numbers = [3, 1, 4, 1, 5, 9]
numbers.reverse()
print(numbers)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['date', 'apple', 'banana', 'cherry']",
            "['banana', 'cherry', 'apple', 'date']",
            "[9, 5, 1, 4, 1, 3]"
        ]
    );
}

#[test]
fn test_py_list_extend_append_insert() {
    let src = r#"
lst = [1, 2]
lst.append([3, 4])
print(lst)

lst2 = [1, 2]
lst2.extend([3, 4])
print(lst2)

lst2.insert(0, 0)
print(lst2)
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, [3, 4]]", "[1, 2, 3, 4]", "[0, 1, 2, 3, 4]"]
    );
}

#[test]
fn test_py_list_pop_remove_clear() {
    let src = r#"
lst = [10, 20, 30, 20, 40]
print(lst.pop())
print(lst.pop(1))
lst.remove(20)  # removes first occurrence
print(lst)
lst.clear()
print(lst)
"#;
    assert_eq!(run_python(src), vec!["40", "20", "[10, 30]", "[]"]);
}

#[test]
fn test_py_list_repetition_multiplication() {
    let src = r#"
zeros = [0] * 5
print(zeros)

nested = [[]] * 3  # note: shared reference
nested[0].append(1)
print(nested)

# un-shared initialization pattern
unshared = [[] for _ in range(3)]
unshared[0].append(1)
print(unshared)
"#;
    assert_eq!(
        run_python(src),
        vec!["[0, 0, 0, 0, 0]", "[[1], [1], [1]]", "[[1], [], []]"]
    );
}

#[test]
fn test_py_list_index_count() {
    let src = r#"
lst = ["a", "b", "c", "b", "a", "b"]
print(lst.count("b"))
print(lst.index("b"))
print(lst.index("b", 2))  # search starting at index 2
"#;
    assert_eq!(run_python(src), vec!["3", "1", "3"]);
}

#[test]
fn test_py_list_shallow_copy() {
    let src = r#"
original = [1, [2, 3], 4]
c1 = original.copy()
c2 = list(original)
c3 = original[:]

c1[1].append(99)
print(original[1])  # inner list modified in all shallow copies
"#;
    assert_eq!(run_python(src), vec!["[2, 3, 99]"]);
}

#[test]
fn test_py_list_concat_iadd() {
    let src = r#"
a = [1, 2]
b = [3, 4]
c = a + b
print(c)
print(a)  # untouched

a += b
print(a)  # mutated
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, 3, 4]", "[1, 2]", "[1, 2, 3, 4]"]
    );
}

#[test]
fn test_py_list_unpacking_extended_target() {
    let src = r#"
first, *middle, last = [1, 2, 3, 4, 5]
print(first, middle, last)
*head, tail = [10, 20]
print(head, tail)
"#;
    assert_eq!(run_python(src), vec!["1 [2, 3, 4] 5", "[10] 20"]);
}
