use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Copy, Pickle & Marshal Serialization — copy, deepcopy, pickle, marshal, __getstate__, __setstate__
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_copy_shallow_vs_deepcopy_behavior() {
    let src = r#"
import copy

original = {"a": [1, 2], "b": {"x": 10}}
shallow = copy.copy(original)
deep = copy.deepcopy(original)

shallow["a"].append(3)
deep["b"]["x"] = 99

print(original["a"])       # modified in shallow
print(original["b"]["x"])  # unmodified in deep
"#;
    assert_eq!(run_python(src), vec!["[1, 2, 3]", "10"]);
}

#[test]
fn test_py_pickle_dumps_loads_object_graph() {
    let src = r#"
import pickle

data = {"users": [{"name": "Alice", "id": 1}], "active": True}
serialized = pickle.dumps(data)
restored = pickle.loads(serialized)

print(restored["users"][0]["name"])
print(restored == data)
"#;
    assert_eq!(run_python(src), vec!["Alice", "True"]);
}

#[test]
fn test_py_pickle_custom_getstate_setstate() {
    let src = r#"
import pickle

class Person:
    def __init__(self, name, secret):
        self.name = name
        self.secret = secret

    def __getstate__(self):
        # Exclude secret from serialized state
        state = self.__dict__.copy()
        del state["secret"]
        return state

    def __setstate__(self, state):
        self.__dict__.update(state)
        self.secret = "default_secret"

p = Person("Alice", "topsecret")
serialized = pickle.dumps(p)
restored = pickle.loads(serialized)

print(restored.name)
print(restored.secret)
"#;
    assert_eq!(run_python(src), vec!["Alice", "default_secret"]);
}

#[test]
fn test_py_marshal_code_object_serialization() {
    let src = r#"
import marshal

code_str = "x = 10 + 20"
code_obj = compile(code_str, "<string>", "exec")
serialized = marshal.dumps(code_obj)
restored_code = marshal.loads(serialized)

scope = {}
exec(restored_code, scope)
print(scope["x"])
"#;
    assert_eq!(run_python(src), vec!["30"]);
}

#[test]
fn test_py_pickle_protocol_highest_supported() {
    let src = r#"
import pickle

data = list(range(100))
serialized = pickle.dumps(data, protocol=pickle.HIGHEST_PROTOCOL)
restored = pickle.loads(serialized)
print(restored == data)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_copy_custom_copy_dunder() {
    let src = r#"
import copy

class CustomCopyable:
    def __init__(self, val):
        self.val = val

    def __copy__(self):
        return CustomCopyable(self.val * 2)

c = CustomCopyable(5)
c_copy = copy.copy(c)
print(c_copy.val)
"#;
    assert_eq!(run_python(src), vec!["10"]);
}

#[test]
fn test_py_copy_custom_deepcopy_memo_dunder() {
    let src = r#"
import copy

class Node:
    def __init__(self, val):
        self.val = val
        self.ref = None

    def __deepcopy__(self, memo):
        if id(self) in memo:
            return memo[id(self)]
        new_node = Node(self.val)
        memo[id(self)] = new_node
        if self.ref:
            new_node.ref = copy.deepcopy(self.ref, memo)
        return new_node

n1 = Node(1)
n2 = Node(2)
n1.ref = n2
n2.ref = n1  # cycle

deep_n1 = copy.deepcopy(n1)
print(deep_n1.val)
print(deep_n1.ref.val)
print(deep_n1.ref.ref is deep_n1)  # cycle preserved
"#;
    assert_eq!(run_python(src), vec!["1", "2", "True"]);
}

#[test]
fn test_py_pickle_reduce_custom_reconstructor() {
    let src = r#"
import pickle

def make_item(name, price):
    return {"name": name, "price": price}

class ItemWrapper:
    def __init__(self, name, price):
        self.name = name
        self.price = price

    def __reduce__(self):
        return (make_item, (self.name, self.price))

iw = ItemWrapper("widget", 10)
serialized = pickle.dumps(iw)
restored = pickle.loads(serialized)
print(restored)
"#;
    assert_eq!(run_python(src), vec!["{'name': 'widget', 'price': 10}"]);
}

#[test]
fn test_py_marshal_primitive_data_types() {
    let src = r#"
import marshal

data = (1, 2.5, "hello", True, [1, 2], {"a": 1})
serialized = marshal.dumps(data)
restored = marshal.loads(serialized)
print(restored == data)
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_pickle_exceptions_handling() {
    let src = r#"
import pickle

class Unpicklable:
    def __getstate__(self):
        raise TypeError("Cannot pickle this class")

u = Unpicklable()
try:
    pickle.dumps(u)
except TypeError as e:
    print("TypeError caught")
"#;
    assert_eq!(run_python(src), vec!["TypeError caught"]);
}
