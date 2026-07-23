use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Collections Counter, Deque & ChainMap — Counter arithmetic, deque rotate, ChainMap scopes, defaultdict
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_counter_elements_most_common() {
    let src = r#"
from collections import Counter

c = Counter(a=3, b=1, c=2)
print(sorted(list(c.elements())))
print(c.most_common(2))
"#;
    assert_eq!(
        run_python(src),
        vec!["['a', 'a', 'a', 'b', 'c', 'c']", "[('a', 3), ('c', 2)]"]
    );
}

#[test]
fn test_py_counter_set_algebra_operators() {
    let src = r#"
from collections import Counter

c1 = Counter(a=3, b=1)
c2 = Counter(a=1, b=2)

print(c1 + c2)
print(c1 - c2)
print(c1 & c2)  # min(c1[x], c2[x])
print(c1 | c2)  # max(c1[x], c2[x])
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Counter({'a': 4, 'b': 3})",
            "Counter({'a': 2})",
            "Counter({'a': 1, 'b': 1})",
            "Counter({'a': 3, 'b': 2})"
        ]
    );
}

#[test]
fn test_py_deque_bounded_rotation() {
    let src = r#"
from collections import deque

dq = deque([1, 2, 3, 4, 5], maxlen=5)
dq.rotate(2)
print(list(dq))
dq.append(99)  # evicts 4 from front
print(list(dq))
"#;
    assert_eq!(run_python(src), vec!["[4, 5, 1, 2, 3]", "[5, 1, 2, 3, 99]"]);
}

#[test]
fn test_py_chainmap_parents_new_child() {
    let src = r#"
from collections import ChainMap

c1 = {"a": 1}
c2 = {"b": 2}
cm = ChainMap(c1, c2)

child = cm.new_child({"a": 10, "c": 3})
print(child["a"])
print(child["b"])
print(child.parents["a"])
"#;
    assert_eq!(run_python(src), vec!["10", "2", "1"]);
}

#[test]
fn test_py_defaultdict_int_list_set_factories() {
    let src = r#"
from collections import defaultdict

d_int = defaultdict(int)
d_list = defaultdict(list)
d_set = defaultdict(set)

d_int["counter"] += 1
d_list["items"].append("item1")
d_set["unique"].add("val1")

print(d_int["counter"])
print(d_list["items"])
print(sorted(list(d_set["unique"])))
"#;
    assert_eq!(run_python(src), vec!["1", "['item1']", "['val1']"]);
}

#[test]
fn test_py_ordereddict_popitem_last_false() {
    let src = r#"
from collections import OrderedDict

od = OrderedDict([("a", 1), ("b", 2), ("c", 3)])
first_key, first_val = od.popitem(last=False)
print(first_key, first_val)
print(list(od.keys()))
"#;
    assert_eq!(run_python(src), vec!["a 1", "['b', 'c']"]);
}

#[test]
fn test_py_deque_extend_extendleft() {
    let src = r#"
from collections import deque

dq = deque([10])
dq.extend([20, 30])
print(list(dq))

dq.extendleft([1, 2])  # note: 1 added then 2 added at left
print(list(dq))
"#;
    assert_eq!(run_python(src), vec!["[10, 20, 30]", "[2, 1, 10, 20, 30]"]);
}

#[test]
fn test_py_counter_total_method_py310() {
    let src = r#"
import sys
from collections import Counter

c = Counter(a=10, b=20, c=30)
if sys.version_info >= (3, 10):
    print(c.total())
else:
    print(sum(c.values()))
"#;
    assert_eq!(run_python(src), vec!["60"]);
}

#[test]
fn test_py_counter_subtract_method() {
    let src = r#"
from collections import Counter

c = Counter(a=3, b=1)
c.subtract(Counter(a=1, b=2))
print(c["a"])
print(c["b"])  # negative counts allowed in subtract!
"#;
    assert_eq!(run_python(src), vec!["2", "-1"]);
}

#[test]
fn test_py_chainmap_maps_list_mutation() {
    let src = r#"
from collections import ChainMap

m1 = {"x": 1}
m2 = {"y": 2}
cm = ChainMap(m1, m2)
print(len(cm.maps))
m1["x"] = 100
print(cm["x"])
"#;
    assert_eq!(run_python(src), vec!["2", "100"]);
}
