use super::helpers::run_js;

// ── WeakMap basic operations ──────────────────────────────
#[test]
fn weakmap_set_and_get() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
const key = {};
wm.set(key, "value");
console.log(wm.get(key));
"#), vec!["value"]);
}

#[test]
fn weakmap_has_returns_correct() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
const k1 = {};
const k2 = {};
wm.set(k1, 1);
console.log(wm.has(k1));
console.log(wm.has(k2));
"#), vec!["true", "false"]);
}

#[test]
fn weakmap_delete_removes_entry() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
const key = {};
wm.set(key, "hello");
wm.delete(key);
console.log(wm.has(key));
"#), vec!["false"]);
}

#[test]
fn weakmap_only_accepts_object_keys() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
try {
  wm.set("string", 1);
  console.log("no error");
} catch (e) {
  console.log("error");
}
"#), vec!["error"]);
}

#[test]
fn weakmap_multiple_keys() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
const k1 = {}, k2 = {}, k3 = {};
wm.set(k1, 1);
wm.set(k2, 2);
wm.set(k3, 3);
console.log(wm.get(k1) + wm.get(k2) + wm.get(k3));
"#), vec!["6"]);
}

#[test]
fn weakmap_private_data_pattern() {
    assert_eq!(run_js(r#"
const _private = new WeakMap();
class Counter {
  constructor() { _private.set(this, { count: 0 }); }
  increment() { _private.get(this).count++; }
  get value() { return _private.get(this).count; }
}
const c = new Counter();
c.increment();
c.increment();
c.increment();
console.log(c.value);
"#), vec!["3"]);
}

#[test]
fn weakmap_returns_undefined_for_missing() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
const key = {};
console.log(wm.get(key) === undefined);
"#), vec!["true"]);
}

#[test]
fn weakmap_constructor_accepts_iterable() {
    assert_eq!(run_js(r#"
const k1 = {}, k2 = {};
const wm = new WeakMap([[k1, "a"], [k2, "b"]]);
console.log(wm.get(k1));
console.log(wm.get(k2));
"#), vec!["a", "b"]);
}

// ── WeakSet basic operations ──────────────────────────────
#[test]
fn weakset_add_and_has() {
    assert_eq!(run_js(r#"
const ws = new WeakSet();
const obj = {};
ws.add(obj);
console.log(ws.has(obj));
"#), vec!["true"]);
}

#[test]
fn weakset_delete_removes() {
    assert_eq!(run_js(r#"
const ws = new WeakSet();
const obj = {};
ws.add(obj);
ws.delete(obj);
console.log(ws.has(obj));
"#), vec!["false"]);
}

#[test]
fn weakset_only_accepts_objects() {
    assert_eq!(run_js(r#"
const ws = new WeakSet();
try {
  ws.add(42);
  console.log("no error");
} catch (e) {
  console.log("error");
}
"#), vec!["error"]);
}

#[test]
fn weakset_constructor_accepts_iterable() {
    assert_eq!(run_js(r#"
const o1 = {}, o2 = {}, o3 = {};
const ws = new WeakSet([o1, o2, o3]);
console.log(ws.has(o1));
console.log(ws.has(o2));
"#), vec!["true", "true"]);
}

#[test]
fn weakset_seen_objects_dedup_pattern() {
    assert_eq!(run_js(r#"
const seen = new WeakSet();
function process(obj) {
  if (seen.has(obj)) return "duplicate";
  seen.add(obj);
  return "new";
}
const a = {};
console.log(process(a));
console.log(process(a));
console.log(process({}));
"#), vec!["new", "duplicate", "new"]);
}

// ── WeakRef ───────────────────────────────────────────────
#[test]
fn weakref_deref_returns_object() {
    assert_eq!(run_js(r#"
let obj = { value: 42 };
const ref = new WeakRef(obj);
const deref = ref.deref();
console.log(deref !== undefined);
console.log(deref.value);
"#), vec!["true", "42"]);
}

#[test]
fn weakref_deref_method_exists() {
    assert_eq!(run_js(r#"
const obj = { x: 1 };
const wr = new WeakRef(obj);
console.log(typeof wr.deref);
"#), vec!["function"]);
}

#[test]
fn weakref_target_not_collected_while_reachable() {
    assert_eq!(run_js(r#"
const target = { id: 99 };
const ref = new WeakRef(target);
const derefed = ref.deref();
console.log(derefed.id);
"#), vec!["99"]);
}

// ── FinalizationRegistry ──────────────────────────────────
#[test]
fn finalizationregistry_constructor_exists() {
    assert_eq!(run_js(r#"
const registry = new FinalizationRegistry(val => {});
console.log(typeof registry.register);
console.log(typeof registry.unregister);
"#), vec!["function", "function"]);
}

#[test]
fn finalizationregistry_register_object() {
    assert_eq!(run_js(r#"
let collected = false;
const registry = new FinalizationRegistry(val => { collected = true; });
let obj = { data: "hello" };
registry.register(obj, "cleanup token");
console.log(typeof obj.data);
"#), vec!["string"]);
}

#[test]
fn finalizationregistry_unregister_returns_bool() {
    assert_eq!(run_js(r#"
const registry = new FinalizationRegistry(() => {});
let obj = {};
const token = {};
registry.register(obj, "value", token);
const removed = registry.unregister(token);
console.log(removed);
"#), vec!["true"]);
}

// ── WeakMap vs Map comparison ─────────────────────────────
#[test]
fn weakmap_is_not_iterable() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
console.log(typeof wm[Symbol.iterator]);
"#), vec!["undefined"]);
}

#[test]
fn weakset_is_not_iterable() {
    assert_eq!(run_js(r#"
const ws = new WeakSet();
console.log(typeof ws[Symbol.iterator]);
"#), vec!["undefined"]);
}

#[test]
fn weakmap_no_size_property() {
    assert_eq!(run_js(r#"
const wm = new WeakMap();
console.log(wm.size === undefined);
"#), vec!["true"]);
}
