use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: list/dict/set/tuple — comprehensions, operations, methods, sorting, slicing, unpacking
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_list_comprehension_basic_and_filtered() {
    let src = r#"
squares = [x ** 2 for x in range(6)]
print(squares)

evens = [x for x in range(10) if x % 2 == 0]
print(evens)

nested = [[i * j for j in range(1, 4)] for i in range(1, 4)]
print(nested)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[0, 1, 4, 9, 16, 25]",
            "[0, 2, 4, 6, 8]",
            "[[1, 2, 3], [2, 4, 6], [3, 6, 9]]"
        ]
    );
}

#[test]
fn test_py_dict_comprehension() {
    let src = r#"
squares = {x: x**2 for x in range(5)}
print(squares)

inverted = {v: k for k, v in squares.items()}
print(inverted)

filtered = {k: v for k, v in squares.items() if v > 4}
print(filtered)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "{0: 0, 1: 1, 2: 4, 3: 9, 4: 16}",
            "{0: 0, 1: 1, 4: 2, 9: 3, 16: 4}",
            "{3: 9, 4: 16}"
        ]
    );
}

#[test]
fn test_py_set_comprehension_and_operations() {
    let src = r#"
a = {1, 2, 3, 4}
b = {3, 4, 5, 6}
print(sorted(a | b))    # union
print(sorted(a & b))    # intersection
print(sorted(a - b))    # difference
print(sorted(a ^ b))    # symmetric difference

squares = {x**2 for x in range(6)}
print(sorted(squares))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[1, 2, 3, 4, 5, 6]",
            "[3, 4]",
            "[1, 2]",
            "[1, 2, 5, 6]",
            "[0, 1, 4, 9, 16, 25]"
        ]
    );
}

#[test]
fn test_py_list_methods_comprehensive() {
    let src = r#"
lst = [3, 1, 4, 1, 5, 9, 2, 6]
print(lst.count(1))
print(lst.index(5))
lst.sort()
print(lst)
lst.reverse()
print(lst[:3])
lst.insert(0, 0)
lst.append(99)
print(lst[0], lst[-1])
"#;
    assert_eq!(
        run_python(src),
        vec!["2", "4", "[1, 1, 2, 3, 4, 5, 6, 9]", "[9, 6, 5]", "0 99"]
    );
}

#[test]
fn test_py_dict_methods_and_access_patterns() {
    let src = r#"
d = {"a": 1, "b": 2, "c": 3}
print(d.get("a"))
print(d.get("z", 99))
d.setdefault("d", 4)
d.setdefault("a", 100)  # won't change existing
print(d["a"], d["d"])
print(sorted(d.keys()))
print(sorted(d.values()))
"#;
    assert_eq!(
        run_python(src),
        vec!["1", "99", "1 4", "['a', 'b', 'c', 'd']", "[1, 2, 3, 4]"]
    );
}

#[test]
fn test_py_dict_merge_operators_py39() {
    let src = r#"
import sys

d1 = {"a": 1, "b": 2}
d2 = {"b": 99, "c": 3}

if sys.version_info >= (3, 9):
    merged = d1 | d2
    print(merged)
    d1 |= d2
    print(d1)
else:
    print({**d1, **d2})
    d1.update(d2)
    print(d1)
"#;
    assert_eq!(
        run_python(src),
        vec!["{'a': 1, 'b': 99, 'c': 3}", "{'a': 1, 'b': 99, 'c': 3}"]
    );
}

#[test]
fn test_py_tuple_packing_unpacking() {
    let src = r#"
t = 1, 2, 3  # packing
print(t)

a, b, c = t  # unpacking
print(a, b, c)

first, *rest = [1, 2, 3, 4, 5]
print(first, rest)

*init, last = [1, 2, 3, 4, 5]
print(init, last)
"#;
    assert_eq!(
        run_python(src),
        vec!["(1, 2, 3)", "1 2 3", "1 [2, 3, 4, 5]", "[1, 2, 3, 4] 5"]
    );
}

#[test]
fn test_py_list_sorting_with_key() {
    let src = r#"
words = ["banana", "apple", "cherry", "kiwi"]
print(sorted(words))
print(sorted(words, key=len))
print(sorted(words, key=lambda w: w[-1]))

people = [("Alice", 30), ("Bob", 25), ("Charlie", 35)]
print(sorted(people, key=lambda p: p[1]))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['apple', 'banana', 'cherry', 'kiwi']",
            "['kiwi', 'apple', 'banana', 'cherry']",
            "['banana', 'apple', 'kiwi', 'cherry']",
            "[('Bob', 25), ('Alice', 30), ('Charlie', 35)]"
        ]
    );
}

#[test]
fn test_py_list_slicing_advanced() {
    let src = r#"
lst = list(range(10))
print(lst[2:7])
print(lst[::3])
print(lst[::-1])
lst[2:5] = [20, 30, 40]
print(lst)
del lst[1:3]
print(lst)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[2, 3, 4, 5, 6]",
            "[0, 3, 6, 9]",
            "[9, 8, 7, 6, 5, 4, 3, 2, 1, 0]",
            "[0, 1, 20, 30, 40, 5, 6, 7, 8, 9]",
            "[0, 30, 40, 5, 6, 7, 8, 9]"
        ]
    );
}

#[test]
fn test_py_set_operations_methods() {
    let src = r#"
s = {1, 2, 3, 4}
s.add(5)
s.discard(3)
s.discard(99)  # no error if not present
print(sorted(s))

try:
    s.remove(99)
except KeyError:
    print("KeyError on remove")

s2 = s.copy()
s2.clear()
print(len(s2))
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, 4, 5]", "KeyError on remove", "0"]
    );
}

#[test]
fn test_py_dict_pop_update_items_iteration() {
    let src = r#"
d = {"a": 1, "b": 2, "c": 3}
val = d.pop("b")
print(val)
d.update({"d": 4, "e": 5})
print(sorted(d.items()))

popped_default = d.pop("z", "missing")
print(popped_default)
"#;
    assert_eq!(
        run_python(src),
        vec!["2", "[('a', 1), ('c', 3), ('d', 4), ('e', 5)]", "missing"]
    );
}

#[test]
fn test_py_frozenset_immutable_set() {
    let src = r#"
fs = frozenset([1, 2, 3, 4])
print(3 in fs)
print(sorted(fs & {2, 3, 5}))
try:
    fs.add(5)
except AttributeError:
    print("AttributeError: frozenset has no add")
d = {fs: "frozen_key"}
print(d[frozenset([1, 2, 3, 4])])
"#;
    assert_eq!(
        run_python(src),
        vec![
            "True",
            "[2, 3]",
            "AttributeError: frozenset has no add",
            "frozen_key"
        ]
    );
}

#[test]
fn test_py_list_extend_concatenate_multiply() {
    let src = r#"
a = [1, 2, 3]
a.extend([4, 5])
print(a)

b = a + [6, 7]
print(b)

c = [0] * 4
print(c)

d = [[]] * 3  # WARNING: shared reference
d[0].append(1)
print(d)  # all three modified
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[1, 2, 3, 4, 5]",
            "[1, 2, 3, 4, 5, 6, 7]",
            "[0, 0, 0, 0]",
            "[[1], [1], [1]]"
        ]
    );
}

#[test]
fn test_py_dict_fromkeys_and_copy() {
    let src = r#"
keys = ["x", "y", "z"]
d = dict.fromkeys(keys, 0)
print(d)

d2 = d.copy()
d2["x"] = 99
print(d["x"])   # shallow copy: original unchanged for immutable vals
print(d2["x"])
"#;
    assert_eq!(run_python(src), vec!["{'x': 0, 'y': 0, 'z': 0}", "0", "99"]);
}
