use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: collections — defaultdict, Counter, deque, OrderedDict, ChainMap, namedtuple, UserDict
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_collections_defaultdict_auto_creates_missing() {
    let src = r#"
from collections import defaultdict

dd = defaultdict(list)
dd["fruits"].append("apple")
dd["fruits"].append("banana")
dd["vegs"].append("carrot")
print(dict(dd))
print(dd["new_key"])  # auto-creates empty list
"#;
    assert_eq!(
        run_python(src),
        vec!["{'fruits': ['apple', 'banana'], 'vegs': ['carrot']}", "[]"]
    );
}

#[test]
fn test_py_collections_defaultdict_nested() {
    let src = r#"
from collections import defaultdict

nested = defaultdict(lambda: defaultdict(int))
nested["users"]["alice"] += 1
nested["users"]["bob"] += 3
print(nested["users"]["alice"])
print(nested["users"]["bob"])
print(nested["users"]["nonexistent"])
"#;
    assert_eq!(run_python(src), vec!["1", "3", "0"]);
}

#[test]
fn test_py_collections_counter_basic() {
    let src = r#"
from collections import Counter

c = Counter("aabbccda")
print(c.most_common(2))
print(c['a'])
print(c['z'])  # missing key returns 0
"#;
    assert_eq!(run_python(src), vec!["[('a', 3), ('b', 2)]", "3", "0"]);
}

#[test]
fn test_py_collections_counter_arithmetic() {
    let src = r#"
from collections import Counter

c1 = Counter(a=3, b=2, c=1)
c2 = Counter(a=1, b=2, d=4)
print(sorted((c1 + c2).items()))
print(sorted((c1 - c2).items()))
print(sorted((c1 & c2).items()))
print(sorted((c1 | c2).items()))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "[('a', 4), ('b', 4), ('c', 1), ('d', 4)]",
            "[('a', 2), ('c', 1)]",
            "[('a', 1), ('b', 2)]",
            "[('a', 3), ('b', 2), ('c', 1), ('d', 4)]"
        ]
    );
}

#[test]
fn test_py_collections_deque_operations() {
    let src = r#"
from collections import deque

d = deque([1, 2, 3])
d.appendleft(0)
d.append(4)
print(list(d))
d.popleft()
d.pop()
print(list(d))
"#;
    assert_eq!(run_python(src), vec!["[0, 1, 2, 3, 4]", "[1, 2, 3]"]);
}

#[test]
fn test_py_collections_deque_maxlen() {
    let src = r#"
from collections import deque

d = deque(maxlen=3)
for i in range(6):
    d.append(i)
print(list(d))  # keeps only last 3
d.appendleft(99)
print(list(d))  # oldest drops off other end
"#;
    assert_eq!(run_python(src), vec!["[3, 4, 5]", "[99, 3, 4]"]);
}

#[test]
fn test_py_collections_deque_rotate() {
    let src = r#"
from collections import deque

d = deque([1, 2, 3, 4, 5])
d.rotate(2)
print(list(d))
d.rotate(-1)
print(list(d))
"#;
    assert_eq!(run_python(src), vec!["[4, 5, 1, 2, 3]", "[5, 1, 2, 3, 4]"]);
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
fn test_py_collections_ordereddict_popitem() {
    let src = r#"
from collections import OrderedDict

od = OrderedDict([("x", 10), ("y", 20), ("z", 30)])
print(od.popitem(last=True))   # LIFO
print(od.popitem(last=False))  # FIFO
"#;
    assert_eq!(run_python(src), vec!["('z', 30)", "('x', 10)"]);
}

#[test]
fn test_py_collections_chainmap_scope_lookup() {
    let src = r#"
from collections import ChainMap

defaults = {"color": "blue", "size": "medium"}
overrides = {"color": "red", "weight": "heavy"}
merged = ChainMap(overrides, defaults)
print(merged["color"])   # from overrides
print(merged["size"])    # from defaults
print(merged["weight"])  # from overrides
"#;
    assert_eq!(run_python(src), vec!["red", "medium", "heavy"]);
}

#[test]
fn test_py_collections_chainmap_new_child() {
    let src = r#"
from collections import ChainMap

base = ChainMap({"a": 1, "b": 2})
child = base.new_child({"a": 99})
print(child["a"])  # child shadows base
print(child["b"])  # falls through to base
print(child.parents["a"])  # access base
"#;
    assert_eq!(run_python(src), vec!["99", "2", "1"]);
}

#[test]
fn test_py_collections_namedtuple_access_and_methods() {
    let src = r#"
from collections import namedtuple

Point = namedtuple("Point", ["x", "y"])
p = Point(3, 4)
print(p.x, p.y)
print(p._asdict())
print(p._replace(x=10))
print(Point._fields)
"#;
    assert_eq!(
        run_python(src),
        vec!["3 4", "{'x': 3, 'y': 4}", "Point(x=10, y=4)", "('x', 'y')"]
    );
}

#[test]
fn test_py_collections_namedtuple_defaults() {
    let src = r#"
from collections import namedtuple

Config = namedtuple("Config", ["host", "port", "debug"], defaults=["localhost", 8080, False])
c1 = Config()
c2 = Config(debug=True)
print(c1)
print(c2.debug)
"#;
    assert_eq!(
        run_python(src),
        vec!["Config(host='localhost', port=8080, debug=False)", "True"]
    );
}

#[test]
fn test_py_collections_counter_elements_and_update() {
    let src = r#"
from collections import Counter

c = Counter(a=3, b=1)
c.update({"b": 2, "c": 5})
print(c["b"])
print(sorted(c.elements()))
"#;
    assert_eq!(
        run_python(src),
        vec![
            "3",
            "['a', 'a', 'a', 'b', 'b', 'b', 'c', 'c', 'c', 'c', 'c']"
        ]
    );
}

#[test]
fn test_py_collections_userdict() {
    let src = r#"
from collections import UserDict

class LowerCaseDict(UserDict):
    def __setitem__(self, key, value):
        super().__setitem__(key.lower(), value)

    def __getitem__(self, key):
        return super().__getitem__(key.lower())

d = LowerCaseDict()
d["Name"] = "Alice"
print(d["name"])
print(d["NAME"])
print(list(d.keys()))
"#;
    assert_eq!(run_python(src), vec!["Alice", "Alice", "['name']"]);
}

#[test]
fn test_py_collections_userlist() {
    let src = r#"
from collections import UserList

class UniqueList(UserList):
    def append(self, item):
        if item not in self.data:
            super().append(item)

ul = UniqueList([1, 2, 3])
ul.append(2)
ul.append(4)
print(ul.data)
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3, 4]"]);
}
