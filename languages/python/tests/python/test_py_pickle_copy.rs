use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: pickle + copy — pickling, deepcopy, shallow copy, __reduce__, __getstate__, __setstate__, copy protocol
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_pickle_basic_roundtrip() {
    let src = r#"
import pickle

data = {"name": "Alice", "scores": [95, 87, 92], "active": True}
serialized = pickle.dumps(data)
restored = pickle.loads(serialized)
print(restored["name"])
print(restored["scores"])
print(restored["active"])
"#;
    assert_eq!(run_python(src), vec!["Alice", "[95, 87, 92]", "True"]);
}

#[test]
fn test_py_pickle_custom_class() {
    let src = r#"
import pickle

class Point:
    def __init__(self, x, y):
        self.x = x
        self.y = y

    def __repr__(self):
        return f"Point({self.x}, {self.y})"

p = Point(3, 4)
data = pickle.dumps(p)
p2 = pickle.loads(data)
print(p2)
print(p2.x, p2.y)
print(p is not p2)
"#;
    assert_eq!(run_python(src), vec!["Point(3, 4)", "3 4", "True"]);
}

#[test]
fn test_py_pickle_getstate_setstate() {
    let src = r#"
import pickle

class Stateful:
    def __init__(self, x, computed=None):
        self.x = x
        self.computed = x * 2  # don't serialize this

    def __getstate__(self):
        return {"x": self.x}  # only save x

    def __setstate__(self, state):
        self.x = state["x"]
        self.computed = self.x * 2  # recompute

s = Stateful(10)
data = pickle.dumps(s)
s2 = pickle.loads(data)
print(s2.x)
print(s2.computed)
"#;
    assert_eq!(run_python(src), vec!["10", "20"]);
}

#[test]
fn test_py_pickle_protocol_versions() {
    let src = r#"
import pickle

data = [1, 2, 3]
for proto in range(pickle.HIGHEST_PROTOCOL + 1):
    serialized = pickle.dumps(data, protocol=proto)
    restored = pickle.loads(serialized)
    print(restored == data)
"#;
    let _proto_count = 6; // protocols 0-5 exist in Python 3.8+
    let result = run_python(src);
    assert!(
        result.iter().all(|r| r == "True"),
        "All protocols should round-trip correctly"
    );
}

#[test]
fn test_py_copy_shallow_vs_deep() {
    let src = r#"
import copy

original = {"key": [1, 2, 3], "nested": {"inner": 42}}
shallow = copy.copy(original)
deep = copy.deepcopy(original)

shallow["key"].append(99)
print(original["key"])  # modified! (shared reference)
print(shallow["key"])

deep["nested"]["inner"] = 99
print(original["nested"]["inner"])  # NOT modified (deep copy)
print(deep["nested"]["inner"])
"#;
    assert_eq!(
        run_python(src),
        vec!["[1, 2, 3, 99]", "[1, 2, 3, 99]", "42", "99"]
    );
}

#[test]
fn test_py_copy_deepcopy_custom_class() {
    let src = r#"
import copy

class Tree:
    def __init__(self, val, children=None):
        self.val = val
        self.children = children or []

root = Tree(1, [Tree(2), Tree(3, [Tree(4)])])
deep = copy.deepcopy(root)

deep.children[0].val = 99
print(root.children[0].val)  # still 2 — not shared
print(deep.children[0].val)
"#;
    assert_eq!(run_python(src), vec!["2", "99"]);
}

#[test]
fn test_py_copy_copy_protocol() {
    let src = r#"
import copy

class Config:
    def __init__(self, host, port):
        self.host = host
        self.port = port

    def __copy__(self):
        return Config(self.host + "_copy", self.port)

    def __deepcopy__(self, memo):
        return Config("deep_" + self.host, self.port + 1000)

cfg = Config("localhost", 8080)
shallow = copy.copy(cfg)
deep = copy.deepcopy(cfg)

print(shallow.host, shallow.port)
print(deep.host, deep.port)
"#;
    assert_eq!(
        run_python(src),
        vec!["localhost_copy 8080", "deep_localhost 9080"]
    );
}

#[test]
fn test_py_pickle_list_of_objects() {
    let src = r#"
import pickle

class Product:
    def __init__(self, name, price):
        self.name = name
        self.price = price

products = [Product("apple", 1.5), Product("banana", 0.8)]
data = pickle.dumps(products)
restored = pickle.loads(data)

print(len(restored))
print(restored[0].name, restored[0].price)
print(restored[1].name, restored[1].price)
"#;
    assert_eq!(run_python(src), vec!["2", "apple 1.5", "banana 0.8"]);
}

#[test]
fn test_py_pickle_file_io() {
    let src = r#"
import pickle, tempfile, os

data = {"a": [1, 2, 3], "b": {"nested": True}}

with tempfile.NamedTemporaryFile(delete=False, suffix=".pkl") as f:
    fname = f.name
    pickle.dump(data, f)

with open(fname, "rb") as f:
    loaded = pickle.load(f)

os.unlink(fname)
print(loaded["a"])
print(loaded["b"])
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "{'nested': True}"]);
}

#[test]
fn test_py_copy_reference_vs_shallow_vs_deep() {
    let src = r#"
import copy

lst = [1, [2, 3], [4, [5, 6]]]
ref = lst
shallow = copy.copy(lst)
deep = copy.deepcopy(lst)

lst[1].append(99)  # modifies inner list

print(ref[1])     # sees the change (same object)
print(shallow[1]) # sees it too (shallow)
print(deep[1])    # does NOT see it
"#;
    assert_eq!(run_python(src), vec!["[2, 3, 99]", "[2, 3, 99]", "[2, 3]"]);
}
