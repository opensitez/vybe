use super::helpers::run_python;

// weakref — ref, proxy, finalize, WeakKeyDictionary, WeakValueDictionary, WeakSet, callbacks, ReferenceError, getweakrefcount, getweakrefs

#[test]
fn test_weakref_ref_liveness_and_callback() {
    let out = run_python(
        r#"
import weakref

class Target: pass

cb_called = []
def callback(reference):
    cb_called.append(True)

obj = Target()
r = weakref.ref(obj, callback)
print(r() is obj)
del obj
print(r() is None)
print(cb_called)
"#,
    );
    assert_eq!(out, vec!["True", "True", "[True]"]);
}

#[test]
fn test_weakref_proxy_attribute_access_and_dereference_error() {
    let out = run_python(
        r#"
import weakref

class Data:
    def __init__(self): self.val = 42

obj = Data()
p = weakref.proxy(obj)
print(p.val)
del obj
try:
    _ = p.val
except ReferenceError:
    print("ReferenceError")
"#,
    );
    assert_eq!(out, vec!["42", "ReferenceError"]);
}

#[test]
fn test_weakref_finalize_cleanup_hook() {
    let out = run_python(
        r#"
import weakref

class Resource: pass

cleaned = []
def cleanup(arg):
    cleaned.append(arg)

r = Resource()
fin = weakref.finalize(r, cleanup, "resource_released")
print(fin.alive)
del r
print(fin.alive)
print(cleaned)
"#,
    );
    assert_eq!(out, vec!["True", "False", "['resource_released']"]);
}

#[test]
fn test_weakref_weak_key_dictionary_auto_removal() {
    let out = run_python(
        r#"
import weakref

class Key: pass

d = weakref.WeakKeyDictionary()
k1 = Key()
k2 = Key()
d[k1] = "val1"
d[k2] = "val2"
print(len(d))
del k1
print(len(d))
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn test_weakref_weak_value_dictionary_auto_removal() {
    let out = run_python(
        r#"
import weakref

class Value: pass

d = weakref.WeakValueDictionary()
v1 = Value()
d["key1"] = v1
print(len(d))
del v1
print(len(d))
"#,
    );
    assert_eq!(out, vec!["1", "0"]);
}

#[test]
fn test_weakref_weak_set_collection() {
    let out = run_python(
        r#"
import weakref

class Item: pass

s = weakref.WeakSet()
i1 = Item()
i2 = Item()
s.add(i1)
s.add(i2)
print(len(s))
del i1
print(len(s))
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn test_weakref_getweakrefcount_and_getweakrefs() {
    let out = run_python(
        r#"
import weakref

class Node: pass

n = Node()
r1 = weakref.ref(n)
r2 = weakref.ref(n)
print(weakref.getweakrefcount(n))
refs = weakref.getweakrefs(n)
print(len(refs))
"#,
    );
    assert_eq!(out, vec!["2", "2"]);
}

#[test]
fn test_weakref_finalize_manual_detach() {
    let out = run_python(
        r#"
import weakref

class Obj: pass

cleaned = []
o = Obj()
fin = weakref.finalize(o, lambda: cleaned.append(1))
fin.detach()
del o
print(cleaned)
"#,
    );
    assert_eq!(out, vec!["[]"]);
}

#[test]
fn test_weakref_finalize_manual_call() {
    let out = run_python(
        r#"
import weakref

class Obj: pass

cleaned = []
o = Obj()
fin = weakref.finalize(o, lambda: cleaned.append("manual"))
fin()
print(cleaned)
print(fin.alive)
"#,
    );
    assert_eq!(out, vec!["['manual']", "False"]);
}

#[test]
fn test_weakref_proxy_callable_object() {
    let out = run_python(
        r#"
import weakref

class CallableObj:
    def __call__(self, x): return x * 2

co = CallableObj()
p = weakref.proxy(co)
print(p(5))
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn test_weakref_ref_hashability_and_equality() {
    let out = run_python(
        r#"
import weakref

class Target: pass

t = Target()
r1 = weakref.ref(t)
r2 = weakref.ref(t)
print(r1 == r2)
print(hash(r1) == hash(r2))
"#,
    );
    assert_eq!(out, vec!["True", "True"]);
}

#[test]
fn test_weakref_builtin_types_unsupported() {
    let out = run_python(
        r#"
import weakref
try:
    weakref.ref(123)
except TypeError:
    print("TypeError")
"#,
    );
    assert_eq!(out, vec!["TypeError"]);
}

#[test]
fn test_weakref_custom_class_with_slots_needs_weakref_slot() {
    let out = run_python(
        r#"
import weakref

class SlottedWithoutWeakref:
    __slots__ = ['x']

class SlottedWithWeakref:
    __slots__ = ['x', '__weakref__']

try:
    weakref.ref(SlottedWithoutWeakref())
except TypeError:
    print("TypeError")

sw = SlottedWithWeakref()
r = weakref.ref(sw)
print(r() is sw)
"#,
    );
    assert_eq!(out, vec!["TypeError", "True"]);
}

#[test]
fn test_weakref_finalize_atexit_parameter() {
    let out = run_python(
        r#"
import weakref

class Dummy: pass

d = Dummy()
fin = weakref.finalize(d, print, "exit")
print(fin.atexit)
fin.atexit = False
print(fin.atexit)
"#,
    );
    assert_eq!(out, vec!["True", "False"]);
}

#[test]
fn test_weakref_weak_value_dict_iteration() {
    let out = run_python(
        r#"
import weakref

class Val:
    def __init__(self, v): self.v = v

d = weakref.WeakValueDictionary()
v1 = Val(1)
v2 = Val(2)
d["a"] = v1
d["b"] = v2
vals = [v.v for v in d.values()]
print(sorted(vals))
"#,
    );
    assert_eq!(out, vec!["[1, 2]"]);
}

#[test]
fn test_weakref_weak_key_dict_iteration() {
    let out = run_python(
        r#"
import weakref

class Key:
    def __init__(self, k): self.k = k

d = weakref.WeakKeyDictionary()
k1 = Key("a")
k2 = Key("b")
d[k1] = 1
d[k2] = 2
keys = [k.k for k in d.keys()]
print(sorted(keys))
"#,
    );
    assert_eq!(out, vec!["['a', 'b']"]);
}

#[test]
fn test_weakref_dead_ref_hashability() {
    let out = run_python(
        r#"
import weakref

class Dummy: pass

d = Dummy()
r = weakref.ref(d)
h1 = hash(r)
del d
try:
    h2 = hash(r)
    print(h1 == h2)
except TypeError:
    print("TypeError")
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_weakref_ref_callback_argument_is_ref() {
    let out = run_python(
        r#"
import weakref

class Obj: pass

ref_received = []
def cb(r):
    ref_received.append(r)

o = Obj()
r = weakref.ref(o, cb)
del o
print(ref_received[0] is r)
"#,
    );
    assert_eq!(out, vec!["True"]);
}

#[test]
fn test_weakref_finalize_peek_result() {
    let out = run_python(
        r#"
import weakref

class Thing: pass

t = Thing()
fin = weakref.finalize(t, int, "123")
res = fin.peek()
print(res)
"#,
    );
    assert_eq!(out, vec!["('123', {})"]);
}

#[test]
fn test_weakref_proxy_equality_with_referent() {
    let out = run_python(
        r#"
import weakref

class Data:
    def __init__(self, val): self.val = val
    def __eq__(self, other): return self.val == getattr(other, "val", None)

d = Data(100)
p = weakref.proxy(d)
print(p == d)
"#,
    );
    assert_eq!(out, vec!["True"]);
}
