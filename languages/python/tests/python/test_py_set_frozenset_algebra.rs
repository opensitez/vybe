use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Set & Frozenset Algebra — union, intersection, difference, symmetric_difference, subsets, frozenset
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_set_operators_vs_methods() {
    let src = r#"
s1 = {1, 2, 3}
s2 = {3, 4, 5}

# Operators require sets
print(sorted(s1 | s2))
print(sorted(s1 & s2))
print(sorted(s1 - s2))
print(sorted(s1 ^ s2))

# Methods accept any iterable
print(sorted(s1.union([3, 4, 5])))
print(sorted(s1.intersection([3, 4])))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[1, 2, 3, 4, 5]",
            "[3]",
            "[1, 2]",
            "[1, 2, 4, 5]",
            "[1, 2, 3, 4, 5]",
            "[3]"
        ]
    );
}

#[test]
fn test_py_set_inplace_mutation_operators() {
    let src = r#"
s = {1, 2}
s |= {3, 4}
print(sorted(s))

s &= {2, 3, 5}
print(sorted(s))

s -= {2}
print(sorted(s))

s ^= {3, 99}
print(sorted(s))
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, 3, 4]", "[2, 3]", "[3]", "[99]"]
    );
}

#[test]
fn test_py_set_subset_superset_disjoint() {
    let src = r#"
a = {1, 2}
b = {1, 2, 3, 4}
c = {5, 6}

print(a.issubset(b))
print(a <= b)
print(a < b)   # proper subset

print(b.issuperset(a))
print(b >= a)
print(b > a)   # proper superset

print(a.isdisjoint(c))
print(a.isdisjoint(b))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "True", "True", "True", "True", "True", "True", "True", "False"
        ]
    );
}

#[test]
fn test_py_set_add_remove_discard_pop() {
    let src = r#"
s = {10, 20}
s.add(30)
s.add(20)  # duplicate ignored
print(sorted(s))

s.remove(20)
print(sorted(s))

s.discard(99)  # no error if missing
try:
    s.remove(99)
except KeyError:
    print("KeyError")

item = s.pop()
print(isinstance(item, int))
"#;
    assert_eq!(
        run_python(src),
        vec!["[10, 20, 30]", "[10, 30]", "KeyError", "True"]
    );
}

#[test]
fn test_py_frozenset_hashable_keys_and_elements() {
    let src = r#"
fs1 = frozenset([1, 2, 3])
fs2 = frozenset([3, 4, 5])

# frozenset can be a dict key or set element
d = {fs1: "group1", fs2: "group2"}
print(d[frozenset([2, 1, 3])])

s = {fs1, fs2}
print(len(s))
"#;
    assert_eq!(run_python(src), vec!["group1", "2"]);
}

#[test]
fn test_py_frozenset_algebra_returns_frozenset() {
    let src = r#"
fs = frozenset([1, 2])
res = fs | {3, 4}
print(type(res).__name__)
print(sorted(res))
"#;
    assert_eq!(run_python(src), vec!["frozenset", "[1, 2, 3, 4]"]);
}

#[test]
fn test_py_set_comprehension_filtering_transform() {
    let src = r#"
words = ["Apple", "banana", "Cherry", "APPLE", "Banana"]
normalized = {w.lower() for w in words}
print(sorted(normalized))
"#;
    assert_eq!(run_python(src), vec!["['apple', 'banana', 'cherry']"]);
}

#[test]
fn test_py_set_update_difference_update() {
    let src = r#"
s = {1, 2, 3, 4, 5}
s.difference_update([1, 3])
print(sorted(s))

s.symmetric_difference_update([4, 6])
print(sorted(s))
"#;
    assert_eq!(run_python(src), vec!["[2, 4, 5]", "[2, 5, 6]"]);
}

#[test]
fn test_py_set_copy_clear() {
    let src = r#"
original = {1, 2, 3}
cp = original.copy()
cp.add(4)
print(sorted(original))
print(sorted(cp))
cp.clear()
print(len(cp))
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "[1, 2, 3, 4]", "0"]);
}

#[test]
fn test_py_empty_set_literals_distinction() {
    let src = r#"
empty_dict = {}
empty_set = set()

print(type(empty_dict).__name__)
print(type(empty_set).__name__)
print(len(empty_set))
"#;
    assert_eq!(run_python(src), vec!["dict", "set", "0"]);
}
