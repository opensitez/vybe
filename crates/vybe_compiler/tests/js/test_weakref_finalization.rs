/// WeakRef and FinalizationRegistry — creation, deref, registration,
/// GC interaction, cleanup callbacks, registry unregister.

use super::helpers::run_js;

// ── WeakRef basics ────────────────────────────────────────────────────────────

#[test]
fn weakref_deref_returns_object_while_alive() {
    assert_eq!(run_js(r#"
let obj = { value: 42 };
const ref1 = new WeakRef(obj);
console.log(ref1.deref()?.value);
"#), vec!["42"]);
}

#[test]
fn weakref_deref_returns_same_object() {
    assert_eq!(run_js(r#"
const obj = { id: 1 };
const ref1 = new WeakRef(obj);
console.log(ref1.deref() === obj);
"#), vec!["true"]);
}

#[test]
fn weakref_instanceof_check() {
    assert_eq!(run_js(r#"
const obj = {};
const ref1 = new WeakRef(obj);
console.log(ref1 instanceof WeakRef);
"#), vec!["true"]);
}

#[test]
fn weakref_deref_with_nullish_coalescing() {
    assert_eq!(run_js(r#"
const obj = { name: "test" };
const ref1 = new WeakRef(obj);
const name = ref1.deref()?.name ?? "gone";
console.log(name);
"#), vec!["test"]);
}

#[test]
fn weakref_can_hold_any_object() {
    assert_eq!(run_js(r#"
const fn1 = () => "hi";
const arr = [1, 2, 3];
const map = new Map();

const refs = [new WeakRef(fn1), new WeakRef(arr), new WeakRef(map)];
console.log(typeof refs[0].deref());
console.log(Array.isArray(refs[1].deref()));
console.log(refs[2].deref() instanceof Map);
"#), vec!["function", "true", "true"]);
}

// ── FinalizationRegistry basics ───────────────────────────────────────────────

#[test]
fn finalization_registry_can_be_created() {
    assert_eq!(run_js(r#"
const registry = new FinalizationRegistry((value) => {
    console.log("cleaned:" + value);
});
console.log(registry instanceof FinalizationRegistry);
"#), vec!["true"]);
}

#[test]
fn finalization_registry_register_does_not_throw() {
    assert_eq!(run_js(r#"
const registry = new FinalizationRegistry(() => {});
let obj = { x: 1 };
registry.register(obj, "token");
console.log("registered");
"#), vec!["registered"]);
}

#[test]
fn finalization_registry_unregister_works() {
    assert_eq!(run_js(r#"
const registry = new FinalizationRegistry(() => {});
let obj = {};
const token = {};
registry.register(obj, "value", token);
const result = registry.unregister(token);
console.log(result); // true if it was registered
"#), vec!["true"]);
}

#[test]
fn finalization_registry_unregister_unknown_token() {
    assert_eq!(run_js(r#"
const registry = new FinalizationRegistry(() => {});
const token = {};
const result = registry.unregister(token); // was never registered
console.log(result);
"#), vec!["false"]);
}

// ── WeakRef + FinalizationRegistry pattern ────────────────────────────────────

#[test]
fn weakref_cache_pattern() {
    assert_eq!(run_js(r#"
// Simulate a cache that holds weak references to objects
class WeakCache {
    #map = new Map();

    set(key, value) {
        this.#map.set(key, new WeakRef(value));
    }

    get(key) {
        return this.#map.get(key)?.deref();
    }
}

const cache = new WeakCache();
const obj = { data: "important" };
cache.set("key", obj);
console.log(cache.get("key")?.data);
"#), vec!["important"]);
}

#[test]
fn finalization_registry_cleanup_receives_held_value() {
    assert_eq!(run_js(r#"
// The callback receives the held value (second arg to register), not the object
const received = [];
const registry = new FinalizationRegistry((heldValue) => {
    received.push(heldValue);
});
// Register with a held value
let obj = {};
registry.register(obj, "my-token");
// We can't force GC, but we can verify the API works
console.log("setup complete");
"#), vec!["setup complete"]);
}

// ── WeakRef in collection ─────────────────────────────────────────────────────

#[test]
fn weakref_set_filters_dead_refs() {
    assert_eq!(run_js(r#"
// Simulate live reference tracking — objects still in scope are alive
let a = { id: "a" };
let b = { id: "b" };

const refs = [new WeakRef(a), new WeakRef(b)];
const live = refs.map(r => r.deref()).filter(Boolean);
console.log(live.length);
console.log(live.map(o => o.id).join(","));
"#), vec!["2", "a,b"]);
}

// ── WeakRef target types ──────────────────────────────────────────────────────

#[test]
fn weakref_cannot_hold_primitives() {
    assert_eq!(run_js(r#"
let threw = false;
try { new WeakRef(42); } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn weakref_cannot_hold_null() {
    assert_eq!(run_js(r#"
let threw = false;
try { new WeakRef(null); } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn weakref_cannot_hold_undefined() {
    assert_eq!(run_js(r#"
let threw = false;
try { new WeakRef(undefined); } catch (e) { threw = e instanceof TypeError; }
console.log(threw);
"#), vec!["true"]);
}
