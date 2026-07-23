use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Special Collections — defaultdict, Counter, OrderedDict, ChainMap, deque, UserDict, UserList
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_collections_defaultdict_custom_factory() {
    let src = r#"
from collections import defaultdict

dd = defaultdict(lambda: "N/A")
dd["name"] = "Alice"
print(dd["name"])
print(dd["missing_key"])
"#;
    assert_eq!(run_python(src), vec!["Alice", "N/A"]);
}

#[test]
fn test_py_collections_counter_most_common() {
    let src = r#"
from collections import Counter

c = Counter("abracadabra")
print(c.most_common(2))
print(c["a"])
print(c["z"])
"#;
    assert_eq!(run_python(src), vec!["[('a', 5), ('b', 2)]", "5", "0"]);
}

#[test]
fn test_py_collections_ordereddict_move_to_end() {
    let src = r#"
from collections import OrderedDict

od = OrderedDict([("a", 1), ("b", 2), ("c", 3)])
od.move_to_end("a")
print(list(od.keys()))

od.move_to_end("c", last=False)
print(list(od.keys()))
"#;
    assert_eq!(run_python(src), vec!["['b', 'c', 'a']", "['c', 'b', 'a']"]);
}

#[test]
fn test_py_collections_chainmap_stacked_scopes() {
    let src = r#"
from collections import ChainMap

defaults = {"color": "red", "user": "guest"}
local_ctx = {"color": "blue"}

cm = ChainMap(local_ctx, defaults)
print(cm["color"])  # local shadows default
print(cm["user"])   # falls back to default

# Modifying mutated primary map
cm["user"] = "admin"
print(local_ctx["user"])
print(defaults["user"])
"#;
    assert_eq!(run_python(src), vec!["blue", "guest", "admin", "guest"]);
}

#[test]
fn test_py_collections_deque_bounded_maxlen() {
    let src = r#"
from collections import deque

dq = deque(maxlen=3)
for i in range(5):
    dq.append(i)

print(list(dq))
dq.appendleft(99)
print(list(dq))
"#;
    assert_eq!(run_python(src), vec!["[2, 3, 4]", "[99, 2, 3]"]);
}

#[test]
fn test_py_collections_userdict_wrapper() {
    let src = r#"
from collections import UserDict

class LowerDict(UserDict):
    def __setitem__(self, key, value):
        super().__setitem__(key.lower(), value)

d = LowerDict()
d["NAME"] = "Alice"
print(d["name"])
print(dict(d))
"#;
    assert_eq!(run_python(src), vec!["Alice", "{'name': 'Alice'}"]);
}

#[test]
fn test_py_collections_userlist_wrapper() {
    let src = r#"
from collections import UserList

class NonNegativeList(UserList):
    def append(self, val):
        if val >= 0:
            super().append(val)

ul = NonNegativeList([1, 2])
ul.append(-5)
ul.append(3)
print(ul.data)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]"]);
}

#[test]
fn test_py_collections_counter_arithmetic_ops() {
    let src = r#"
from collections import Counter

c1 = Counter(a=3, b=1)
c2 = Counter(a=1, b=2)

print(c1 + c2)
print(c1 - c2)  # keeps only positive counts
"#;
    assert_eq!(
        run_python(src),
        vec!["Counter({'a': 4, 'b': 3})", "Counter({'a': 2})"]
    );
}

#[test]
fn test_py_collections_deque_rotate() {
    let src = r#"
from collections import deque

dq = deque([1, 2, 3, 4, 5])
dq.rotate(2)
print(list(dq))

dq.rotate(-1)
print(list(dq))
"#;
    assert_eq!(run_python(src), vec!["[4, 5, 1, 2, 3]", "[5, 1, 2, 3, 4]"]);
}

#[test]
fn test_py_collections_defaultdict_nested_tree() {
    let src = r#"
from collections import defaultdict

def tree():
    return defaultdict(tree)

t = tree()
t["a"]["b"]["c"] = 42
print(t["a"]["b"]["c"])
"#;
    assert_eq!(run_python(src), vec!["42"]);
}
