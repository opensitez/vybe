use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Dictionary Views & Mutation — dict.keys(), values(), items(), view operations, set operations, merging
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_dict_views_reflect_mutation() {
    let src = r#"
d = {"a": 1, "b": 2}
keys = d.keys()
values = d.values()
items = d.items()

print(list(keys))
d["c"] = 3
print(list(keys))     # reflect additions
print(list(values))
print(list(items))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "['a', 'b']",
            "['a', 'b', 'c']",
            "[1, 2, 3]",
            "[('a', 1), ('b', 2), ('c', 3)]"
        ]
    );
}

#[test]
fn test_py_dict_keys_set_operations() {
    let src = r#"
d1 = {"a": 1, "b": 2, "c": 3}
d2 = {"b": 20, "c": 30, "d": 40}

k1 = d1.keys()
k2 = d2.keys()

print(sorted(k1 & k2))  # intersection
print(sorted(k1 | k2))  # union
print(sorted(k1 - k2))  # difference
print(sorted(k1 ^ k2))  # symmetric difference
"#;
    assert_eq!(
        run_python(src),
        vec!["['b', 'c']", "['a', 'b', 'c', 'd']", "['a']", "['a', 'd']"]
    );
}

#[test]
fn test_py_dict_items_set_operations() {
    let src = r#"
d1 = {"a": 1, "b": 2}
d2 = {"b": 2, "c": 3}

items1 = d1.items()
items2 = d2.items()

print(sorted(items1 & items2))
print(sorted(items1 | items2))
"#;
    assert_eq!(
        run_python(src),
        vec!["[('b', 2)]", "[('a', 1), ('b', 2), ('c', 3)]"]
    );
}

#[test]
fn test_py_dict_update_variants() {
    let src = r#"
d = {"a": 1}
d.update({"b": 2}, c=3)
print(sorted(d.items()))

d.update([("d", 4), ("e", 5)])
print(sorted(d.items()))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[('a', 1), ('b', 2), ('c', 3)]",
            "[('a', 1), ('b', 2), ('c', 3), ('d', 4), ('e', 5)]"
        ]
    );
}

#[test]
fn test_py_dict_pop_popitem_clear() {
    let src = r#"
d = {"x": 10, "y": 20, "z": 30}
print(d.pop("y"))
print(d.pop("nonexistent", 999))
k, v = d.popitem()
print(k, v)
d.clear()
print(len(d))
"#;
    assert_eq!(run_python(src), vec!["20", "999", "z 30", "0"]);
}

#[test]
fn test_py_dict_setdefault_behavior() {
    let src = r#"
d = {"a": [1]}
d.setdefault("a", []).append(2)
d.setdefault("b", []).append(10)
print(d["a"])
print(d["b"])
"#;
    assert_eq!(run_python(src), vec!["[1, 2]", "[10]"]);
}

#[test]
fn test_py_dict_merge_union_operators() {
    let src = r#"
d1 = {"a": 1, "b": 2}
d2 = {"b": 20, "c": 30}

merged = d1 | d2
print(merged)
print(d1)  # d1 untouched

d1 |= d2
print(d1)  # d1 mutated in-place
"#;
    assert_eq!(
        run_python(src),
        vec![
            "{'a': 1, 'b': 20, 'c': 30}",
            "{'a': 1, 'b': 2}",
            "{'a': 1, 'b': 20, 'c': 30}"
        ]
    );
}

#[test]
fn test_py_dict_comprehension_transformations() {
    let src = r#"
scores = {"alice": 85, "bob": 92, "charlie": 78}
passed = {k.capitalize(): v for k, v in scores.items() if v >= 80}
print(sorted(passed.items()))
"#;
    assert_eq!(run_python(src), vec!["[('Alice', 85), ('Bob', 92)]"]);
}

#[test]
fn test_py_dict_missing_dunder_override() {
    let src = r#"
class DefaultDictCustom(dict):
    def __missing__(self, key):
        self[key] = f"default_{key}"
        return self[key]

d = DefaultDictCustom()
print(d["foo"])
print(d["bar"])
print(d)
"#;
    assert_eq!(
        run_python(src),
        vec![
            "default_foo",
            "default_bar",
            "{'foo': 'default_foo', 'bar': 'default_bar'}"
        ]
    );
}

#[test]
fn test_py_dict_insertion_order_preservation() {
    let src = r#"
d = {}
d["z"] = 1
d["a"] = 2
d["m"] = 3
print(list(d.keys()))
d.pop("a")
d["a"] = 99
print(list(d.keys()))  # 'a' moves to end upon re-insertion
"#;
    assert_eq!(run_python(src), vec!["['z', 'a', 'm']", "['z', 'm', 'a']"]);
}
