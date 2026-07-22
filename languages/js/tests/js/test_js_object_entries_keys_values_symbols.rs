use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Object Reflection (`Object.keys()`, `values()`, `entries()`, `getOwnPropertySymbols()`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_object_keys_returns_own_enumerable_string_keys() {
    let src = r#"
const obj = { b: 2, a: 1 };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
console.log(Object.keys(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["b,a"]);
}

#[test]
fn test_js_object_values_returns_own_enumerable_string_values() {
    let src = r#"
const obj = { b: 20, a: 10 };
Object.defineProperty(obj, "hidden", { value: 30, enumerable: false });
console.log(Object.values(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["20,10"]);
}

#[test]
fn test_js_object_entries_returns_key_value_pairs() {
    let src = r#"
const obj = { x: 1, y: 2 };
const pairs = Object.entries(obj);
console.log(pairs.map(p => p.join("=")).join("|"));
"#;
    assert_eq!(run_js(src), vec!["x=1|y=2"]);
}

#[test]
fn test_js_object_from_entries_reconstructs_object() {
    let src = r#"
const entries = [["a", 10], ["b", 20]];
const obj = Object.fromEntries(entries);
console.log(`${obj.a}:${obj.b}`);
"#;
    assert_eq!(run_js(src), vec!["10:20"]);
}

#[test]
fn test_js_object_from_entries_map_source() {
    let src = r#"
const map = new Map([["key1", "val1"], ["key2", "val2"]]);
const obj = Object.fromEntries(map);
console.log(`${obj.key1}:${obj.key2}`);
"#;
    assert_eq!(run_js(src), vec!["val1:val2"]);
}

#[test]
fn test_js_object_get_own_property_symbols_reflects_symbols_only() {
    let src = r#"
const s1 = Symbol("s1");
const s2 = Symbol("s2");
const obj = { stringKey: 1, [s1]: "val1", [s2]: "val2" };
const symbols = Object.getOwnPropertySymbols(obj);
console.log(symbols.length + "|" + (symbols[0] === s1) + "|" + (symbols[1] === s2));
"#;
    assert_eq!(run_js(src), vec!["2|true|true"]);
}

#[test]
fn test_js_object_get_own_property_symbols_includes_non_enumerable() {
    let src = r#"
const s = Symbol("hiddenSym");
const obj = {};
Object.defineProperty(obj, s, { value: "secret", enumerable: false });
const symbols = Object.getOwnPropertySymbols(obj);
console.log(symbols.length + "|" + obj[symbols[0]]);
"#;
    assert_eq!(run_js(src), vec!["1|secret"]);
}

#[test]
fn test_js_reflect_own_keys_combines_string_and_symbol_keys() {
    let src = r#"
const s = Symbol("sym");
const obj = { b: 2, 1: "num", [s]: "symVal" };
Object.defineProperty(obj, "hidden", { value: 3, enumerable: false });
const keys = Reflect.ownKeys(obj);
console.log(keys.map(String).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,b,hidden,Symbol(sym)"]); // Integer indices, string keys, symbol keys!
}

#[test]
fn test_js_object_keys_coerces_primitive_argument_to_object() {
    let src = r#"
console.log(Object.keys("hi").join(",") + "|" + Object.keys(123).length);
"#;
    assert_eq!(run_js(src), vec!["0,1|0"]);
}

#[test]
fn test_js_object_keys_null_or_undefined_throws_typeerror() {
    let src = r#"
try {
    Object.keys(null);
} catch (e) {
    console.log("Object.keys Null TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Object.keys Null TypeError"]);
}

#[test]
fn test_js_object_values_evaluates_getters() {
    let src = r#"
const obj = {
    get a() { return "GetterA"; }
};
console.log(Object.values(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["GetterA"]);
}

#[test]
fn test_js_object_entries_evaluates_getters() {
    let src = r#"
const obj = {
    get a() { return 100; }
};
console.log(Object.entries(obj)[0].join("="));
"#;
    assert_eq!(run_js(src), vec!["a=100"]);
}

#[test]
fn test_js_object_from_entries_iterable_yields_pair_arrays() {
    let src = r#"
function* pairGen() {
    yield ["x", 1];
    yield ["y", 2];
}
const obj = Object.fromEntries(pairGen());
console.log(`${obj.x}:${obj.y}`);
"#;
    assert_eq!(run_js(src), vec!["1:2"]);
}

#[test]
fn test_js_object_from_entries_invalid_pair_element_throws_typeerror() {
    let src = r#"
try {
    Object.fromEntries([["valid", 1], "invalid_pair"]);
} catch (e) {
    console.log("fromEntries Invalid Pair TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["fromEntries Invalid Pair TypeError"]);
}

#[test]
fn test_js_object_keys_ignores_prototype_chain() {
    let src = r#"
const proto = { protoKey: 1 };
const obj = Object.create(proto);
obj.ownKey = 2;
console.log(Object.keys(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["ownKey"]);
}

#[test]
fn test_js_object_values_sparse_array_holes_ignored() {
    let src = r#"
const sparse = [1, , 3];
console.log(Object.values(sparse).join(",")); // Sparse holes are non-enumerable, omitted from Object.values!
"#;
    assert_eq!(run_js(src), vec!["1,3"]);
}

#[test]
fn test_js_object_entries_sparse_array_holes_ignored() {
    let src = r#"
const sparse = [10, , 30];
const entries = Object.entries(sparse);
console.log(entries.map(e => e.join("=")).join("|"));
"#;
    assert_eq!(run_js(src), vec!["0=10|2=30"]);
}

#[test]
fn test_js_object_keys_integer_indices_sorted_first() {
    let src = r#"
const obj = { 10: "ten", 2: "two", "b": "B", 1: "one", "a": "A" };
console.log(Object.keys(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["1,2,10,b,a"]);
}

#[test]
fn test_js_object_get_own_property_names_includes_non_enumerable() {
    let src = r#"
const obj = { a: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
console.log(Object.getOwnPropertyNames(obj).join(","));
"#;
    assert_eq!(run_js(src), vec!["a,hidden"]);
}

#[test]
fn test_js_object_from_entries_symbol_keys() {
    let src = r#"
const s = Symbol("key");
const obj = Object.fromEntries([[s, "symbolVal"]]);
console.log(obj[s]);
"#;
    assert_eq!(run_js(src), vec!["symbolVal"]);
}
