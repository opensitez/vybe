use super::helpers::run_python;

// ═══════════════════════════════════════════════════════════
// Python: weakref + garbage collection — weakref.ref, WeakValueDictionary, WeakKeyDictionary, gc module, __del__, finalize
// ═══════════════════════════════════════════════════════════

#[test]
fn test_py_weakref_basic_reference() {
    let src = r#"
import weakref

class Resource:
    def __init__(self, name):
        self.name = name

r = Resource("DB")
wr = weakref.ref(r)

print(wr() is r)      # alive
print(wr().name)

del r
import gc; gc.collect()
print(wr() is None)   # dead
"#;
    assert_eq!(run_python(src), vec!["True", "DB", "True"]);
}

#[test]
fn test_py_weakref_callback() {
    let src = r#"
import weakref

log = []

class Obj:
    def __init__(self, name):
        self.name = name

def on_destroy(ref):
    log.append("destroyed")

o = Obj("temp")
wr = weakref.ref(o, on_destroy)

del o
import gc; gc.collect()
print(log)
print(wr() is None)
"#;
    assert_eq!(run_python(src), vec!["['destroyed']", "True"]);
}

#[test]
fn test_py_weakref_weakvalue_dictionary() {
    let src = r#"
import weakref, gc

class Session:
    def __init__(self, sid):
        self.sid = sid

cache = weakref.WeakValueDictionary()
s1 = Session("abc")
s2 = Session("xyz")
cache["abc"] = s1
cache["xyz"] = s2

print("abc" in cache)
del s1
gc.collect()
print("abc" in cache)  # cleaned up automatically
print("xyz" in cache)  # still alive
"#;
    assert_eq!(run_python(src), vec!["True", "False", "True"]);
}

#[test]
fn test_py_weakref_weakkey_dictionary() {
    let src = r#"
import weakref, gc

class Key:
    def __init__(self, val):
        self.val = val

    def __hash__(self):
        return hash(self.val)

    def __eq__(self, other):
        return self.val == other.val

d = weakref.WeakKeyDictionary()
k1 = Key("a")
d[k1] = "value_a"
print(d[k1])

del k1
gc.collect()
print(len(d))  # entry removed
"#;
    assert_eq!(run_python(src), vec!["value_a", "0"]);
}

#[test]
fn test_py_weakref_finalize() {
    let src = r#"
import weakref

log = []

class Resource:
    pass

r = Resource()
weakref.finalize(r, log.append, "finalized")

del r
import gc; gc.collect()
print(log)
"#;
    assert_eq!(run_python(src), vec!["['finalized']"]);
}

#[test]
fn test_py_gc_collect_and_get_stats() {
    let src = r#"
import gc

gc.collect()  # run collection
stats = gc.get_stats()
print(len(stats))  # 3 generations
print(all("collections" in s for s in stats))
print(gc.isenabled())
"#;
    assert_eq!(run_python(src), vec!["3", "True", "True"]);
}

#[test]
fn test_py_gc_reference_cycle_detection() {
    let src = r#"
import gc

class Node:
    def __init__(self):
        self.ref = None

gc.collect()
before = gc.get_count()[0]

# Create a cycle
a = Node()
b = Node()
a.ref = b
b.ref = a
del a, b

# gc should detect and clean the cycle
collected = gc.collect()
print(collected >= 2)  # at least 2 objects collected
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_gc_get_referrers() {
    let src = r#"
import gc

class Tracked:
    pass

obj = Tracked()
container = [obj]

refs = gc.get_referrers(obj)
print(any(r is container for r in refs))
"#;
    assert_eq!(run_python(src), vec!["True"]);
}

#[test]
fn test_py_del_method_cleanup() {
    let src = r#"
log = []

class Connection:
    def __init__(self, url):
        self.url = url
        log.append(f"open:{url}")

    def __del__(self):
        log.append(f"close:{self.url}")

def use_connection():
    conn = Connection("db://localhost")
    return "done"

result = use_connection()
import gc; gc.collect()
print(result)
print(log)
"#;
    assert_eq!(
        run_python(src),
        vec!["done", "['open:db://localhost', 'close:db://localhost']"]
    );
}

#[test]
fn test_py_weakref_proxy() {
    let src = r#"
import weakref

class Widget:
    def __init__(self, name):
        self.name = name

    def display(self):
        return f"Widget: {self.name}"

w = Widget("button")
proxy = weakref.proxy(w)

print(proxy.name)
print(proxy.display())
del w
import gc; gc.collect()
try:
    proxy.name
except ReferenceError:
    print("ReferenceError: dead proxy")
"#;
    assert_eq!(
        run_python(src),
        vec!["button", "Widget: button", "ReferenceError: dead proxy"]
    );
}
