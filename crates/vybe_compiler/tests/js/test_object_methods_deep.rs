/// Object methods deep dive — Object.create, Object.assign, Object.keys/values/entries,
/// Object.fromEntries, Object.hasOwn, Object.getPrototypeOf/setPrototypeOf,
/// Object.getOwnPropertyNames/Symbols, Object.defineProperties, Object.freeze/seal deep.
use super::helpers::run_js;

// ── Object.create ─────────────────────────────────────────────────────────────

#[test]
fn object_create_with_prototype() {
    assert_eq!(
        run_js(
            r#"
const proto = { greet() { return "hello from " + this.name; } };
const obj = Object.create(proto);
obj.name = "World";
console.log(obj.greet());
console.log(Object.getPrototypeOf(obj) === proto);
"#
        ),
        vec!["hello from World", "true"]
    );
}

#[test]
fn object_create_null_prototype() {
    assert_eq!(
        run_js(
            r#"
const pure = Object.create(null);
pure.x = 1;
console.log(Object.getPrototypeOf(pure));
console.log(pure.x);
"#
        ),
        vec!["null", "1"]
    );
}

#[test]
fn object_create_with_property_descriptors() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create({}, {
    x: { value: 10, writable: true, enumerable: true, configurable: true },
    y: { value: 20, writable: false, enumerable: true, configurable: true }
});
console.log(obj.x);
console.log(obj.y);
"#
        ),
        vec!["10", "20"]
    );
}

// ── Object.assign ─────────────────────────────────────────────────────────────

#[test]
fn object_assign_merges_multiple_sources() {
    assert_eq!(
        run_js(
            r#"
const target = { a: 1 };
const result = Object.assign(target, { b: 2 }, { c: 3 }, { b: 99 });
console.log(result === target);
console.log(result.a + "," + result.b + "," + result.c);
"#
        ),
        vec!["true", "1,99,3"]
    );
}

#[test]
fn object_assign_only_copies_enumerable_own() {
    assert_eq!(
        run_js(
            r#"
const src = Object.create({ inherited: true });
Object.defineProperty(src, "hidden", { value: 1, enumerable: false });
src.visible = 2;

const target = Object.assign({}, src);
console.log("inherited" in target);
console.log("hidden" in target);
console.log(target.visible);
"#
        ),
        vec!["false", "false", "2"]
    );
}

// ── Object.keys/values/entries ────────────────────────────────────────────────

#[test]
fn object_keys_only_own_enumerable() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Object.keys(obj);
console.log(keys.sort().join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn object_values_returns_values() {
    assert_eq!(
        run_js(
            r#"
const obj = { x: 10, y: 20, z: 30 };
const vals = Object.values(obj);
console.log(vals.sort((a,b) => a-b).join(","));
"#
        ),
        vec!["10,20,30"]
    );
}

#[test]
fn object_entries_returns_key_value_pairs() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2 };
const entries = Object.entries(obj);
console.log(entries.map(([k,v]) => k+"="+v).sort().join(","));
"#
        ),
        vec!["a=1,b=2"]
    );
}

// ── Object.fromEntries ────────────────────────────────────────────────────────

#[test]
fn object_from_entries_basic() {
    assert_eq!(
        run_js(
            r#"
const entries = [["a", 1], ["b", 2], ["c", 3]];
const obj = Object.fromEntries(entries);
console.log(obj.a + "," + obj.b + "," + obj.c);
"#
        ),
        vec!["1,2,3"]
    );
}

#[test]
fn object_from_entries_from_map() {
    assert_eq!(
        run_js(
            r#"
const map = new Map([["key1", "val1"], ["key2", "val2"]]);
const obj = Object.fromEntries(map);
console.log(obj.key1);
console.log(obj.key2);
"#
        ),
        vec!["val1", "val2"]
    );
}

#[test]
fn object_from_entries_transform_pattern() {
    assert_eq!(
        run_js(
            r#"
const obj = { a: 1, b: 2, c: 3 };
const doubled = Object.fromEntries(
    Object.entries(obj).map(([k, v]) => [k, v * 2])
);
console.log(doubled.a + "," + doubled.b + "," + doubled.c);
"#
        ),
        vec!["2,4,6"]
    );
}

// ── Object.hasOwn ─────────────────────────────────────────────────────────────

#[test]
fn object_has_own_vs_prototype() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create({ inherited: true });
obj.own = true;
console.log(Object.hasOwn(obj, "own"));
console.log(Object.hasOwn(obj, "inherited"));
"#
        ),
        vec!["true", "false"]
    );
}

#[test]
fn object_has_own_on_null_prototype_object() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.create(null);
obj.x = 1;
console.log(Object.hasOwn(obj, "x"));
// This works even though obj.hasOwnProperty is undefined
"#
        ),
        vec!["true"]
    );
}

// ── getOwnPropertyNames/Symbols ───────────────────────────────────────────────

#[test]
fn get_own_property_names_includes_non_enumerable() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
obj.a = 1;
Object.defineProperty(obj, "b", { value: 2, enumerable: false });
const names = Object.getOwnPropertyNames(obj).sort();
console.log(names.join(","));
"#
        ),
        vec!["a,b"]
    );
}

#[test]
fn get_own_property_symbols_finds_symbol_keys() {
    assert_eq!(
        run_js(
            r#"
const sym = Symbol("test");
const obj = { [sym]: "value", normal: 1 };
const symbols = Object.getOwnPropertySymbols(obj);
console.log(symbols.length);
console.log(obj[symbols[0]]);
"#
        ),
        vec!["1", "value"]
    );
}

// ── Object.freeze / Object.seal deep ─────────────────────────────────────────

#[test]
fn freeze_prevents_adding_properties() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.freeze({ a: 1 });
obj.b = 2; // silently fails in non-strict
console.log("b" in obj);
"#
        ),
        vec!["false"]
    );
}

#[test]
fn freeze_prevents_modifying_properties() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.freeze({ a: 1 });
obj.a = 99;
console.log(obj.a);
"#
        ),
        vec!["1"]
    );
}

#[test]
fn seal_prevents_adding_but_allows_modifying() {
    assert_eq!(
        run_js(
            r#"
const obj = Object.seal({ x: 1 });
obj.y = 2; // adding fails
obj.x = 99; // modifying works
console.log("y" in obj);
console.log(obj.x);
"#
        ),
        vec!["false", "99"]
    );
}

#[test]
fn is_frozen_is_sealed_checks() {
    assert_eq!(
        run_js(
            r#"
const frozen = Object.freeze({});
const sealed = Object.seal({});
const plain = {};
console.log(Object.isFrozen(frozen));
console.log(Object.isSealed(sealed));
console.log(Object.isFrozen(plain));
console.log(Object.isSealed(plain));
"#
        ),
        vec!["true", "true", "false", "false"]
    );
}

// ── Object.defineProperties ───────────────────────────────────────────────────

#[test]
fn object_define_properties_batch() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
Object.defineProperties(obj, {
    x: { value: 10, writable: true, enumerable: true, configurable: true },
    y: { value: 20, writable: true, enumerable: true, configurable: true },
});
obj.z = obj.x + obj.y;
console.log(obj.x);
console.log(obj.y);
console.log(obj.z);
"#
        ),
        vec!["10", "20", "30"]
    );
}

// ── Object.getOwnPropertyDescriptor ──────────────────────────────────────────

#[test]
fn get_own_property_descriptor_accessor() {
    assert_eq!(
        run_js(
            r#"
const obj = {};
let _val = 0;
Object.defineProperty(obj, "v", {
    get() { return _val; },
    set(x) { _val = x; },
    enumerable: true,
    configurable: true
});
const desc = Object.getOwnPropertyDescriptor(obj, "v");
console.log(typeof desc.get);
console.log(typeof desc.set);
console.log("value" in desc);
"#
        ),
        vec!["function", "function", "false"]
    );
}
