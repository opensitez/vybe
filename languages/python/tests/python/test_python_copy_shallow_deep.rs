use super::helpers::run_python;

// copy — copy, deepcopy, __copy__, __deepcopy__, memo dict handling cyclic structures, custom objects, containers

#[test]
fn test_copy_shallow_copy_list_references() {
    let out = run_python(
        r#"
import copy
inner = [1, 2]
orig = [inner, 3]
c = copy.copy(orig)
print(c is not orig)
print(c[0] is orig[0])
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_deepcopy_list_isolation() {
    let out = run_python(
        r#"
import copy
inner = [1, 2]
orig = [inner, 3]
dc = copy.deepcopy(orig)
print(dc is not orig)
print(dc[0] is not orig[0])
dc[0].append(99)
print(orig[0])
print(dc[0])
"#,
    );
    assert_eq!(out, vec!["True", "True", "[1, 2]", "[1, 2, 99]"]);
}

#[test]
fn test_copy_deepcopy_handles_cyclic_references() {
    let out = run_python(
        r#"
import copy
a = []
b = [a]
a.append(b)

dc_a = copy.deepcopy(a)
print(dc_a[0][0] is dc_a)
print(dc_a is not a)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_custom_class_copy_hook() {
    let out = run_python(
        r#"
import copy

class Custom:
    def __init__(self, val):
        self.val = val
    def __copy__(self):
        return Custom(self.val * 2)

obj = Custom(10)
c = copy.copy(obj)
print(c.val)
"#,
    );
    assert_eq!(out, vec!["20"]);
}

#[test]
fn test_copy_custom_class_deepcopy_hook_with_memo() {
    let out = run_python(
        r#"
import copy

class CustomDeep:
    def __init__(self, data):
        self.data = data
    def __deepcopy__(self, memo):
        copied_data = copy.deepcopy(self.data, memo)
        return CustomDeep(copied_data)

cd = CustomDeep([1, [2, 3]])
dc = copy.deepcopy(cd)
print(dc.data)
print(dc.data[1] is not cd.data[1])
"#,
    );
    assert_eq!(out, vec!["[1, [2, 3]]", "True"]);
}

#[test]
fn test_copy_dict_shallow_and_deepcopy() {
    let out = run_python(
        r#"
import copy
d = {"a": [1, 2], "b": 3}
c = copy.copy(d)
dc = copy.deepcopy(d)

d["a"].append(3)
print(c["a"])
print(dc["a"])
"#,
    );
    assert_eq!(out, vec!["[1, 2, 3]", "[1, 2]"]);
}

#[test]
fn test_copy_set_and_frozenset() {
    let out = run_python(
        r#"
import copy
s = {1, 2, 3}
fs = frozenset([4, 5])
c_s = copy.copy(s)
dc_fs = copy.deepcopy(fs)
print(c_s == s and c_s is not s)
print(dc_fs is fs)  # Frozenset is immutable, returned as-is
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_tuple_immutable_optimization() {
    let out = run_python(
        r#"
import copy
t = (1, 2, 3)
c_t = copy.copy(t)
print(c_t is t)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_copy_tuple_with_mutable_contents_deepcopy() {
    let out = run_python(
        r#"
import copy
t = ([1], [2])
dc_t = copy.deepcopy(t)
print(dc_t is not t)
print(dc_t[0] is not t[0])
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_atomic_types_immutable_identity() {
    let out = run_python(
        r#"
import copy
x = 100
s = "string"
f = 3.14
print(copy.copy(x) is x)
print(copy.copy(s) is s)
print(copy.copy(f) is f)
"#,
    );
    assert_eq!(out, vec!["True", "True", "True"]);
}

#[test]
fn test_copy_deepcopy_memo_reuse() {
    let out = run_python(
        r#"
import copy
shared = [10, 20]
container = [shared, shared]
dc = copy.deepcopy(container)
print(dc[0] is dc[1])
print(dc[0] is not shared)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_class_instance_without_copy_method() {
    let out = run_python(
        r#"
import copy

class Simple:
    def __init__(self, x):
        self.x = x

s = Simple(5)
c = copy.copy(s)
dc = copy.deepcopy(s)
print(c is not s and c.x == 5)
print(dc is not s and dc.x == 5)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_error_for_uncopyable_objects() {
    let out = run_python(
        r#"
import copy, sys
try:
    copy.copy(sys)
except copy.Error:
    print("Error")
"#,
    );
    assert_eq!(out, vec!["Error"]);
}

#[test]
fn test_copy_deepcopy_defaultdict() {
    let out = run_python(
        r#"
from collections import defaultdict
import copy

dd = defaultdict(list)
dd["a"].append(1)
dc = copy.deepcopy(dd)
print(dc["a"])
print(dc.default_factory is list)
"#,
    );
    assert_eq!(out, vec!["[1]", "True"]);
}

#[test]
fn test_copy_deepcopy_ordereddict() {
    let out = run_python(
        r#"
from collections import OrderedDict
import copy

od = OrderedDict([("a", [1]), ("b", [2])])
dc = copy.deepcopy(od)
print(list(dc.keys()))
print(dc["a"] is not od["a"])
"#,
    );
    assert_eq!(out, vec!["['a', 'b']", "True"]);
}

#[test]
fn test_copy_dispatch_table_customization() {
    let out = run_python(
        r#"
import copy

class Special:
    def __init__(self, val): self.val = val

def _copy_special(obj):
    return Special(obj.val + 100)

copy._copy_dispatch[Special] = _copy_special
s = Special(5)
c = copy.copy(s)
print(c.val)
"#,
    );
    assert_eq!(out, vec!["105"]);
}

#[test]
fn test_copy_deepcopy_bytearray() {
    let out = run_python(
        r#"
import copy
b = bytearray(b"hello")
dc = copy.deepcopy(b)
print(dc is not b)
print(dc == b)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_copy_deepcopy_complex_nested_structure() {
    let out = run_python(
        r#"
import copy
data = {"list": [{"a": 1}, {"b": 2}], "tuple": (1, [2, 3])}
dc = copy.deepcopy(data)
dc["list"][0]["a"] = 99
print(data["list"][0]["a"])
"#,
    );
    assert_eq!(out, vec!["1"]);
}

#[test]
fn test_copy_copy_reg_pickle_support() {
    let out = run_python(
        r#"
import copy

class ReductionCopy:
    def __reduce__(self):
        return (ReductionCopy, ())

rc = ReductionCopy()
c = copy.copy(rc)
print(isinstance(c, ReductionCopy))
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_copy_deepcopy_preserves_slots() {
    let out = run_python(
        r#"
import copy

class Slotted:
    __slots__ = ['x', 'y']
    def __init__(self, x, y):
        self.x = x
        self.y = y

s = Slotted([1], [2])
dc = copy.deepcopy(s)
print(dc.x is not s.x)
print(dc.x == s.x)
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}
