use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: Weakref & Reference Tracking — weakref.ref, WeakValueDictionary, WeakKeyDictionary, finalize, proxy
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_weakref_basic_reference_lifecycle() {
    let src = r#"
import weakref, gc

class Target:
    def __init__(self, name):
        self.name = name

obj = Target("live_object")
ref = weakref.ref(obj)

print(ref() is obj)
print(ref().name)

del obj
gc.collect()
print(ref() is None)
"#;
    assert_eq!(run_python(src), vec!["True", "live_object", "True"]);
}

#[test]
fn test_py_weakref_callback_on_garbage_collection() {
    let src = r#"
import weakref, gc

log = []

class Target: pass

def on_finalize(ref):
    log.append("finalized")

obj = Target()
ref = weakref.ref(obj, on_finalize)

del obj
gc.collect()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['finalized']"]);
}

#[test]
fn test_py_weakref_weakvaluedictionary_cleanup() {
    let src = r#"
import weakref, gc

class Node:
    def __init__(self, val):
        self.val = val

cache = weakref.WeakValueDictionary()
n1 = Node(10)
n2 = Node(20)

cache["n1"] = n1
cache["n2"] = n2

print("n1" in cache)
del n1
gc.collect()
print("n1" in cache)
print("n2" in cache)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True"]);
}

#[test]
fn test_py_weakref_weakkeydictionary_cleanup() {
    let src = r#"
import weakref, gc

class Key:
    def __init__(self, name):
        self.name = name

d = weakref.WeakKeyDictionary()
k1 = Key("k1")
d[k1] = "val1"

print(d[k1])
del k1
gc.collect()
print(len(d))
"#;
    assert_eq!(run_python(src), vec!["val1", "0"]);
}

#[test]
fn test_py_weakref_finalize_cleanup_callable() {
    let src = r#"
import weakref, gc

cleaned = []

class Resource:
    pass

r = Resource()
fin = weakref.finalize(r, cleaned.append, "cleaned_up")

print(fin.alive)
del r
gc.collect()
print(fin.alive)
print(cleaned)
"#;
    assert_eq!(run_python(src), vec!["True", "False", "['cleaned_up']"]);
}

#[test]
fn test_py_weakref_proxy_transparent_access() {
    let src = r#"
import weakref, gc

class Person:
    def __init__(self, name):
        self.name = name

p = Person("Alice")
proxy = weakref.proxy(p)
print(proxy.name)

del p
gc.collect()
try:
    _ = proxy.name
except ReferenceError:
    print("ReferenceError: weakly referenced object no longer exists")
"#;
    assert_eq!(
        run_python(src),
        vec![
            "Alice",
            "ReferenceError: weakly referenced object no longer exists"
        ]
    );
}

#[test]
fn test_py_weakref_getweakrefcount_getweakrefs() {
    let src = r#"
import weakref

class Target: pass

t = Target()
print(weakref.getweakrefcount(t))
r1 = weakref.ref(t)
r2 = weakref.ref(t)
print(weakref.getweakrefcount(t))
print(len(weakref.getweakrefs(t)))
"#;
    assert_eq!(run_python(src), vec!["0", "2", "2"]);
}

#[test]
fn test_py_weakref_set_collection() {
    let src = r#"
import weakref, gc

class Obj: pass

o1 = Obj()
o2 = Obj()
ws = weakref.WeakSet([o1, o2])

print(len(ws))
del o1
gc.collect()
print(len(ws))
"#;
    assert_eq!(run_python(src), vec!["2", "1"]);
}

#[test]
fn test_py_weakref_finalize_atexit_option() {
    let src = r#"
import weakref

class Res: pass

r = Res()
fin = weakref.finalize(r, print, "finalized_on_exit")
print(fin.atexit)
fin.atexit = False
print(fin.atexit)
"#;
    assert_eq!(run_python(src), vec!["True", "False"]);
}

#[test]
fn test_py_weakref_cannot_ref_builtin_types() {
    let src = r#"
import weakref

try:
    weakref.ref(42)
except TypeError:
    print("TypeError: cannot create weak reference to 'int'")
"#;
    assert_eq!(
        run_python(src),
        vec!["TypeError: cannot create weak reference to 'int'"]
    );
}
