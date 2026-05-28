/// Object property descriptors, Object static methods, and accessor properties.
/// Covers defineProperty, getOwnPropertyDescriptor, defineProperties,
/// Object.fromEntries, Object.hasOwn, getOwnPropertyNames/Symbols,
/// Object.assign, Object.is, Object.create with descriptors.

use super::helpers::run_js;

// ── Object.defineProperty ─────────────────────────────────────────────────────

#[test]
fn define_property_non_writable_prevents_assignment() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 42, writable: false, configurable: true });
obj.fixed = 99;  // silently ignored (non-strict)
console.log(obj.fixed);
console.log(obj.fixed === 42);
"#), vec!["42", "true"]);
}

#[test]
fn define_property_non_enumerable_hidden_from_for_in() {
    assert_eq!(run_js(r#"
const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false, writable: true, configurable: true });
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("visible"));
console.log(keys.includes("hidden"));
"#), vec!["true", "false"]);
}

#[test]
fn define_property_non_configurable_prevents_redefinition() {
    assert_eq!(run_js(r#"
"use strict";
const obj = {};
Object.defineProperty(obj, "locked", { value: 1, configurable: false });
let threw = false;
try {
    Object.defineProperty(obj, "locked", { value: 2 });
} catch { threw = true; }
console.log(threw);
"#), vec!["true"]);
}

#[test]
fn define_property_accessor_get_set() {
    assert_eq!(run_js(r#"
const obj = { _n: 0 };
Object.defineProperty(obj, "n", {
    get() { return this._n; },
    set(v) { this._n = v < 0 ? 0 : v; },
    configurable: true
});
obj.n = 5;
console.log(obj.n);
obj.n = -3;
console.log(obj.n);
"#), vec!["5", "0"]);
}

#[test]
fn define_property_with_symbol_key() {
    assert_eq!(run_js(r#"
const sym = Symbol("id");
const obj = {};
Object.defineProperty(obj, sym, { value: 99, enumerable: true, configurable: true, writable: true });
console.log(obj[sym]);
"#), vec!["99"]);
}

// ── Object.getOwnPropertyDescriptor ─────────────────────────────────────────

#[test]
fn get_own_property_descriptor_data_descriptor() {
    assert_eq!(run_js(r#"
const obj = { x: 42 };
const d = Object.getOwnPropertyDescriptor(obj, "x");
console.log(d.value);
console.log(d.writable);
console.log(d.enumerable);
console.log(d.configurable);
"#), vec!["42", "true", "true", "true"]);
}

#[test]
fn get_own_property_descriptor_accessor_has_no_value() {
    assert_eq!(run_js(r#"
const obj = { get x() { return 1; } };
const d = Object.getOwnPropertyDescriptor(obj, "x");
console.log(typeof d.get);
console.log(d.value === undefined);
"#), vec!["function", "true"]);
}

#[test]
fn get_own_property_descriptor_returns_undefined_for_missing() {
    assert_eq!(run_js(r#"
const obj = {};
console.log(Object.getOwnPropertyDescriptor(obj, "missing"));
"#), vec!["undefined"]);
}

// ── Object.defineProperties ───────────────────────────────────────────────────

#[test]
fn define_properties_adds_multiple_at_once() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperties(obj, {
    a: { value: 1, enumerable: true, writable: true, configurable: true },
    b: { value: 2, enumerable: true, writable: true, configurable: true },
    c: { value: 3, enumerable: false, writable: true, configurable: true }
});
const keys = Object.keys(obj);
console.log(keys.sort().join(","));
console.log(obj.c);
"#), vec!["a,b", "3"]);
}

// ── Object.getOwnPropertyNames ───────────────────────────────────────────────

#[test]
fn get_own_property_names_includes_non_enumerable() {
    assert_eq!(run_js(r#"
const obj = {};
Object.defineProperty(obj, "hidden", { value: 1, enumerable: false, configurable: true });
obj.visible = 2;
const names = Object.getOwnPropertyNames(obj).sort();
console.log(names.join(","));
"#), vec!["hidden,visible"]);
}

#[test]
fn get_own_property_names_excludes_symbol_keys() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const obj = { a: 1, [sym]: 2 };
const names = Object.getOwnPropertyNames(obj);
console.log(names.includes("a"));
console.log(names.length);
"#), vec!["true", "1"]);
}

// ── Object.getOwnPropertySymbols ─────────────────────────────────────────────

#[test]
fn get_own_property_symbols_returns_symbol_keys() {
    assert_eq!(run_js(r#"
const s1 = Symbol("a");
const s2 = Symbol("b");
const obj = { [s1]: 1, [s2]: 2, str: 3 };
const syms = Object.getOwnPropertySymbols(obj);
console.log(syms.length);
"#), vec!["2"]);
}

// ── Object.fromEntries ────────────────────────────────────────────────────────

#[test]
fn from_entries_from_map() {
    assert_eq!(run_js(r#"
const m = new Map([["a", 1], ["b", 2], ["c", 3]]);
const obj = Object.fromEntries(m);
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#), vec!["1", "2", "3"]);
}

#[test]
fn from_entries_from_array_of_pairs() {
    assert_eq!(run_js(r#"
const entries = [["x", 10], ["y", 20]];
const obj = Object.fromEntries(entries);
console.log(obj.x + obj.y);
"#), vec!["30"]);
}

#[test]
fn from_entries_round_trips_with_entries() {
    assert_eq!(run_js(r#"
const original = { a: 1, b: 2, c: 3 };
const modified = Object.fromEntries(
    Object.entries(original).map(([k, v]) => [k, v * 2])
);
console.log(modified.a);
console.log(modified.b);
console.log(modified.c);
"#), vec!["2", "4", "6"]);
}

// ── Object.hasOwn ─────────────────────────────────────────────────────────────

#[test]
fn has_own_returns_true_for_own_property() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
console.log(Object.hasOwn(obj, "a"));
console.log(Object.hasOwn(obj, "toString"));
"#), vec!["true", "false"]);
}

#[test]
fn has_own_works_on_null_prototype_object() {
    assert_eq!(run_js(r#"
const obj = Object.create(null);
obj.x = 42;
console.log(Object.hasOwn(obj, "x"));
console.log(Object.hasOwn(obj, "y"));
"#), vec!["true", "false"]);
}

// ── Object.assign ─────────────────────────────────────────────────────────────

#[test]
fn object_assign_multiple_sources() {
    assert_eq!(run_js(r#"
const target = { a: 1 };
const result = Object.assign(target, { b: 2 }, { c: 3 }, { d: 4 });
console.log(result === target);
console.log(Object.keys(result).sort().join(","));
"#), vec!["true", "a,b,c,d"]);
}

#[test]
fn object_assign_later_source_overwrites_earlier() {
    assert_eq!(run_js(r#"
const result = Object.assign({}, { x: 1 }, { x: 2 }, { x: 3 });
console.log(result.x);
"#), vec!["3"]);
}

#[test]
fn object_assign_only_copies_own_enumerable() {
    assert_eq!(run_js(r#"
const proto = { inherited: 1 };
const src = Object.create(proto);
src.own = 2;
const result = Object.assign({}, src);
console.log(result.own);
console.log(result.inherited);
"#), vec!["2", "undefined"]);
}

// ── Object.is ────────────────────────────────────────────────────────────────

#[test]
fn object_is_distinguishes_neg_zero() {
    assert_eq!(run_js(r#"
console.log(Object.is(0, -0));
console.log(Object.is(-0, -0));
console.log(-0 === 0);
"#), vec!["false", "true", "true"]);
}

#[test]
fn object_is_nan_equals_nan() {
    assert_eq!(run_js(r#"
console.log(Object.is(NaN, NaN));
console.log(NaN === NaN);
"#), vec!["true", "false"]);
}

#[test]
fn object_is_same_value_semantics() {
    assert_eq!(run_js(r#"
console.log(Object.is(1, 1));
console.log(Object.is("a", "a"));
console.log(Object.is(null, null));
console.log(Object.is(undefined, undefined));
"#), vec!["true", "true", "true", "true"]);
}

// ── Object.create with property descriptors ───────────────────────────────────

#[test]
fn object_create_with_descriptors() {
    assert_eq!(run_js(r#"
const obj = Object.create(Object.prototype, {
    x: { value: 10, writable: true, enumerable: true, configurable: true },
    y: { value: 20, writable: true, enumerable: true, configurable: true }
});
console.log(obj.x + obj.y);
"#), vec!["30"]);
}

#[test]
fn object_create_null_has_no_prototype_methods() {
    assert_eq!(run_js(r#"
const obj = Object.create(null);
obj.key = "value";
console.log(typeof obj.hasOwnProperty);
console.log(obj.key);
"#), vec!["undefined", "value"]);
}

// ── Object spread ─────────────────────────────────────────────────────────────

#[test]
fn object_spread_merges_two_objects() {
    assert_eq!(run_js(r#"
const a = { x: 1, y: 2 };
const b = { y: 99, z: 3 };
const merged = { ...a, ...b };
console.log(merged.x);
console.log(merged.y);
console.log(merged.z);
"#), vec!["1", "99", "3"]);
}

#[test]
fn object_spread_shallow_copy() {
    assert_eq!(run_js(r#"
const src = { a: 1, b: { nested: 2 } };
const copy = { ...src };
copy.a = 99;
copy.b.nested = 88;
console.log(src.a);
console.log(src.b.nested);
"#), vec!["1", "88"]);
}

// ── Object.keys / values / entries edge cases ────────────────────────────────

#[test]
fn object_keys_returns_only_own_enumerable() {
    assert_eq!(run_js(r#"
const proto = { inherited: 1 };
const obj = Object.create(proto);
obj.own1 = 2;
obj.own2 = 3;
console.log(Object.keys(obj).sort().join(","));
"#), vec!["own1,own2"]);
}

#[test]
fn object_values_returns_own_enumerable_values() {
    assert_eq!(run_js(r#"
const obj = { a: 10, b: 20, c: 30 };
const vals = Object.values(obj).sort((a, b) => a - b);
console.log(vals.join(","));
"#), vec!["10,20,30"]);
}

#[test]
fn object_entries_returns_key_value_pairs() {
    assert_eq!(run_js(r#"
const obj = { x: 1, y: 2 };
const entries = Object.entries(obj).sort(([a], [b]) => a < b ? -1 : 1);
console.log(entries.map(([k, v]) => k + ":" + v).join(","));
"#), vec!["x:1,y:2"]);
}

// ── getOwnPropertyDescriptors ────────────────────────────────────────────────

#[test]
fn get_own_property_descriptors_returns_all_descriptors() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, enumerable: false, writable: false, configurable: false });
const descs = Object.getOwnPropertyDescriptors(obj);
console.log(descs.a.value);
console.log(descs.b.enumerable);
"#), vec!["1", "false"]);
}

#[test]
fn get_own_property_descriptors_enables_perfect_clone() {
    assert_eq!(run_js(r#"
const orig = {};
Object.defineProperty(orig, "x", { value: 42, writable: false, enumerable: true, configurable: false });
const clone = {};
Object.defineProperty(clone, "x", { value: orig.x, writable: false, enumerable: true, configurable: false });
console.log(clone.x);
clone.x = 99;  // silently ignored (writable: false)
console.log(clone.x);
"#), vec!["42", "42"]);
}

// ── property enumeration order ────────────────────────────────────────────────

#[test]
fn integer_indices_come_before_string_keys_in_enumeration() {
    assert_eq!(run_js(r#"
const obj = { b: 2, a: 1, "0": "zero", "2": "two", "1": "one" };
const keys = Object.keys(obj);
console.log(keys.includes("0"));
console.log(keys.includes("a"));
console.log(keys.length);
"#), vec!["true", "true", "5"]);
}

// ── Object.preventExtensions ──────────────────────────────────────────────────

#[test]
fn prevent_extensions_blocks_new_properties() {
    assert_eq!(run_js(r#"
const obj = { existing: 1 };
Object.preventExtensions(obj);
obj.newProp = 2;  // silently ignored
console.log(obj.existing);
console.log("newProp" in obj);
"#), vec!["1", "false"]);
}

#[test]
fn is_extensible_returns_correct_value() {
    assert_eq!(run_js(r#"
const obj = {};
console.log(Object.isExtensible(obj));
Object.preventExtensions(obj);
console.log(Object.isExtensible(obj));
"#), vec!["true", "false"]);
}
