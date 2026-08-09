/// Map and Set advanced — Map iteration order, Map.get/has edge cases,
/// Set intersection (manual), WeakMap/WeakSet usage, Map as cache,
/// Set for deduplication, Map/Set composition, size tracking.
use super::helpers::run_js;

// ── Map insertion order ───────────────────────────────────────────────────────

#[test]
fn map_preserves_insertion_order() {
    assert_eq!(
        run_js(
            r#"
const m = new Map();
m.set("c", 3); m.set("a", 1); m.set("b", 2);
const keys = [];
m.forEach((v, k) => keys.push(k));
console.log(keys.join(","));
"#
        ),
        vec!["c,a,b"]
    );
}

#[test]
fn map_iteration_via_for_of() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["x", 10], ["y", 20], ["z", 30]]);
const result = [];
for (const [k, v] of m) result.push(k + "=" + v);
console.log(result.join(","));
"#
        ),
        vec!["x=10,y=20,z=30"]
    );
}

// ── Map with object keys ──────────────────────────────────────────────────────

#[test]
fn map_uses_reference_equality_for_object_keys() {
    assert_eq!(
        run_js(
            r#"
const key1 = { id: 1 };
const key2 = { id: 1 }; // different object, same content
const m = new Map();
m.set(key1, "value1");
console.log(m.has(key1));
console.log(m.has(key2)); // different reference
console.log(m.size);
"#
        ),
        vec!["true", "false", "1"]
    );
}

#[test]
fn map_can_use_any_value_as_key() {
    assert_eq!(
        run_js(
            r#"
const m = new Map();
m.set(null, "null key");
m.set(undefined, "undefined key");
m.set(NaN, "nan key");
m.set(true, "bool key");
console.log(m.get(null));
console.log(m.get(undefined));
console.log(m.get(NaN)); // NaN === NaN in Map (SameValueZero)
console.log(m.get(true));
"#
        ),
        vec!["null key", "undefined key", "nan key", "bool key"]
    );
}

// ── Map update semantics ──────────────────────────────────────────────────────

#[test]
fn map_set_updates_existing_key() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1]]);
m.set("a", 99);
console.log(m.get("a"));
console.log(m.size);
"#
        ),
        vec!["99", "1"]
    );
}

#[test]
fn map_delete_reduces_size() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
m.delete("b");
console.log(m.size);
console.log(m.has("b"));
console.log(m.delete("notExist"));
"#
        ),
        vec!["2", "false", "false"]
    );
}

// ── Map iterators ─────────────────────────────────────────────────────────────

#[test]
fn map_keys_values_entries_iterators() {
    assert_eq!(
        run_js(
            r#"
const m = new Map([["a", 1], ["b", 2]]);
console.log([...m.keys()].join(","));
console.log([...m.values()].join(","));
console.log([...m.entries()].map(([k,v]) => k+"="+v).join(","));
"#
        ),
        vec!["a,b", "1,2", "a=1,b=2"]
    );
}

// ── Set basics and deduplication ──────────────────────────────────────────────

#[test]
fn set_deduplicates_values() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([1, 2, 3, 2, 1, 4]);
console.log(s.size);
console.log([...s].join(","));
"#
        ),
        vec!["4", "1,2,3,4"]
    );
}

#[test]
fn set_nan_deduplication() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([NaN, NaN]);
console.log(s.size); // NaN is deduplicated in Set
"#
        ),
        vec!["1"]
    );
}

#[test]
fn set_operations_manual() {
    assert_eq!(
        run_js(
            r#"
const a = new Set([1, 2, 3, 4]);
const b = new Set([3, 4, 5, 6]);

// Union
const union = new Set([...a, ...b]);
console.log([...union].sort((x,y) => x-y).join(","));

// Intersection
const intersection = new Set([...a].filter(x => b.has(x)));
console.log([...intersection].join(","));

// Difference (a - b)
const diff = new Set([...a].filter(x => !b.has(x)));
console.log([...diff].join(","));
"#
        ),
        vec!["1,2,3,4,5,6", "3,4", "1,2"]
    );
}

// ── WeakMap usage ────────────────────────────────────────────────────────────

#[test]
fn weakmap_stores_per_object_data() {
    assert_eq!(
        run_js(
            r#"
const meta = new WeakMap();
const obj1 = {};
const obj2 = {};
meta.set(obj1, { created: 2024 });
meta.set(obj2, { created: 2025 });
console.log(meta.get(obj1).created);
console.log(meta.get(obj2).created);
console.log(meta.has({})); // different object
"#
        ),
        vec!["2024", "2025", "false"]
    );
}

#[test]
fn weakmap_has_and_delete() {
    assert_eq!(
        run_js(
            r#"
const wm = new WeakMap();
const key = {};
wm.set(key, "value");
console.log(wm.has(key));
wm.delete(key);
console.log(wm.has(key));
"#
        ),
        vec!["true", "false"]
    );
}

// ── WeakSet usage ─────────────────────────────────────────────────────────────

#[test]
fn weakset_tracks_objects() {
    assert_eq!(
        run_js(
            r#"
const seen = new WeakSet();
const a = {};
const b = {};
seen.add(a);
console.log(seen.has(a));
console.log(seen.has(b));
seen.add(b);
seen.delete(a);
console.log(seen.has(a));
"#
        ),
        vec!["true", "false", "false"]
    );
}

#[test]
fn weakset_for_cycle_detection() {
    assert_eq!(
        run_js(
            r#"
function hasCycle(obj, seen = new WeakSet()) {
    if (typeof obj !== "object" || obj === null) return false;
    if (seen.has(obj)) return true;
    seen.add(obj);
    return Object.values(obj).some(v => hasCycle(v, seen));
}

const normal = { a: { b: { c: 1 } } };
console.log(hasCycle(normal));

const cyclic = {};
cyclic.self = cyclic;
console.log(hasCycle(cyclic));
"#
        ),
        vec!["false", "true"]
    );
}

// ── Map as cache / memoization ────────────────────────────────────────────────

#[test]
fn map_as_lru_like_cache() {
    assert_eq!(
        run_js(
            r#"
const cache = new Map();
let computeCount = 0;

function compute(key) {
    if (cache.has(key)) return cache.get(key);
    computeCount++;
    const result = key * key;
    cache.set(key, result);
    return result;
}

compute(5); compute(5); compute(6); compute(5);
console.log(computeCount);  // 2 unique computations
console.log(cache.get(5));
"#
        ),
        vec!["2", "25"]
    );
}

// ── Set from iterables ────────────────────────────────────────────────────────

#[test]
fn set_from_generator() {
    assert_eq!(
        run_js(
            r#"
function* range(n) { for (let i = 0; i < n; i++) yield i; }
const s = new Set(range(5));
console.log(s.size);
console.log([...s].join(","));
"#
        ),
        vec!["5", "0,1,2,3,4"]
    );
}

#[test]
fn set_clear_empties_set() {
    assert_eq!(
        run_js(
            r#"
const s = new Set([1, 2, 3]);
s.clear();
console.log(s.size);
console.log(s.has(1));
"#
        ),
        vec!["0", "false"]
    );
}

#[test]
fn map_foreach_this_arg_binding_context() {
    assert_eq!(
        run_js(
            r#"
const ctx = { factor: 10 };
const m = new Map([["a", 2]]);
m.forEach(function(v, k) {
    console.log(k + ":" + (v * this.factor));
}, ctx);
"#
        ),
        vec!["a:20"]
    );
}
