use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Dictionary Views & Advanced Methods — views set algebra, pop, setdefault, update, mapping ABC, key requirements
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dict_keys_set_algebra() {
    let src = r#"
d1 = {"a": 1, "b": 2, "c": 3}
d2 = {"b": 20, "c": 30, "d": 40}

print(sorted(d1.keys() & d2.keys()))
print(sorted(d1.keys() | d2.keys()))
print(sorted(d1.keys() - d2.keys()))
"#;
    assert_eq!(
        run_python(src),
        vec!["['b', 'c']", "['a', 'b', 'c', 'd']", "['a']"]
    );
}

#[test]
fn test_py_dict_items_set_algebra() {
    let src = r#"
d1 = {"a": 1, "b": 2}
d2 = {"b": 2, "c": 3}

print(sorted(d1.items() & d2.items()))
print(sorted(d1.items() | d2.items()))
"#;
    assert_eq!(
        run_python(src),
        vec!["[('b', 2)]", "[('a', 1), ('b', 2), ('c', 3)]"]
    );
}

#[test]
fn test_py_dict_pop_default_handling() {
    let src = r#"
d = {"x": 10, "y": 20}
val = d.pop("x")
print(val)
fallback = d.pop("z", 999)
print(fallback)
"#;
    assert_eq!(run_python(src), vec!["10", "999"]);
}

#[test]
fn test_py_dict_setdefault_grouping_pattern() {
    let src = r#"
words = ["apple", "ant", "banana", "bear", "cat"]
grouped = {}
for w in words:
    grouped.setdefault(w[0], []).append(w)

print(sorted(grouped["a"]))
print(sorted(grouped["b"]))
"#;
    assert_eq!(
        run_python(src),
        vec!["['ant', 'apple']", "['banana', 'bear']"]
    );
}

#[test]
fn test_py_dict_update_iterable_and_kwargs() {
    let src = r#"
d = {"a": 1}
d.update([("b", 2), ("c", 3)], d=4)
print(sorted(d.items()))
"#;
    assert_eq!(
        run_python(src),
        vec!["[('a', 1), ('b', 2), ('c', 3), ('d', 4)]"]
    );
}

#[test]
fn test_py_dict_key_hashability_requirement() {
    let src = r#"
d = {}
try:
    d[[1, 2]] = "unhashable"
except TypeError as e:
    print("TypeError: unhashable type: 'list'")
"#;
    assert_eq!(run_python(src), vec!["TypeError: unhashable type: 'list'"]);
}

#[test]
fn test_py_dict_tuple_and_frozenset_keys() {
    let src = r#"
d = {}
d[(1, 2)] = "tuple_key"
d[frozenset([3, 4])] = "frozen_key"

print(d[(1, 2)])
print(d[frozenset([4, 3])])
"#;
    assert_eq!(run_python(src), vec!["tuple_key", "frozen_key"]);
}

#[test]
fn test_py_collections_abc_mapping_custom_implementation() {
    let src = r#"
from collections.abc import Mapping

class ConstantMapping(Mapping):
    def __init__(self, default_value):
        self._val = default_value

    def __getitem__(self, key):
        return self._val

    def __len__(self):
        return 1

    def __iter__(self):
        yield "default"

m = ConstantMapping(42)
print(m["foo"])
print(m["bar"])
print("foo" in m)
"#;
    assert_eq!(run_python(src), vec!["42", "42", "True"]);
}

#[test]
fn test_py_dict_comprehension_conditional_filtering() {
    let src = r#"
raw = {"a": 1, "b": 2, "c": 3, "d": 4}
evens = {k: v * 10 for k, v in raw.items() if v % 2 == 0}
print(sorted(evens.items()))
"#;
    assert_eq!(run_python(src), vec!["[('b', 20), ('d', 40)]"]);
}

#[test]
fn test_py_dict_copy_shallow_isolation() {
    let src = r#"
d1 = {"a": 1, "b": 2}
d2 = d1.copy()
d2["c"] = 3
print("c" in d1)
print("c" in d2)
"#;
    assert_eq!(run_python(src), vec!["False", "True"]);
}

#[test]
fn test_py_dict_fromkeys_factory() {
    let src = r#"
keys = ["x", "y", "z"]
d = dict.fromkeys(keys, 0)
print(sorted(d.items()))
"#;
    assert_eq!(run_python(src), vec!["[('x', 0), ('y', 0), ('z', 0)]"]);
}

#[test]
fn test_py_dict_or_operator_merge_py39() {
    let src = r#"
d1 = {"a": 1, "b": 2}
d2 = {"b": 20, "c": 30}
merged = d1 | d2
print(sorted(merged.items()))
"#;
    assert_eq!(run_python(src), vec!["[('a', 1), ('b', 20), ('c', 30)]"]);
}

#[test]
fn test_py_dict_values_view_iteration() {
    let src = r#"
d = {"a": 10, "b": 20, "c": 30}
vals = d.values()
print(sum(vals))
"#;
    assert_eq!(run_python(src), vec!["60"]);
}

#[test]
fn test_py_dict_reversed_keys_py38() {
    let src = r#"
d = {"a": 1, "b": 2, "c": 3}
print(list(reversed(d)))
"#;
    assert_eq!(run_python(src), vec!["['c', 'b', 'a']"]);
}

#[test]
fn test_py_dict_clear_empties_contents() {
    let src = r#"
d = {"a": 1, "b": 2}
d.clear()
print(len(d))
print(bool(d))
"#;
    assert_eq!(run_python(src), vec!["0", "False"]);
}
