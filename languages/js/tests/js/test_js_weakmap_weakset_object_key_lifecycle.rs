use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: WeakMap & WeakSet Object & Symbol Key Collections (ES2023)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_weakmap_set_get_has_delete_flow() {
    let src = r#"
const wm = new WeakMap();
const key = { id: 1 };
wm.set(key, "PrivateData");

console.log(wm.get(key) + "|" + wm.has(key));
wm.delete(key);
console.log(wm.has(key));
"#;
    assert_eq!(run_js(src), vec!["PrivateData|true", "false"]);
}

#[test]
fn test_js_weakmap_primitive_key_throws_typeerror() {
    let src = r#"
const wm = new WeakMap();
try {
    wm.set("string_key", 100);
} catch (e) {
    console.log("WeakMap Primitive Key TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["WeakMap Primitive Key TypeError"]);
}

#[test]
fn test_js_weakmap_symbol_key_support_es2023() {
    let src = r#"
const wm = new WeakMap();
const sym = Symbol("weakKey");
wm.set(sym, "SymbolValue");
console.log(wm.get(sym) + "|" + wm.has(sym));
"#;
    assert_eq!(run_js(src), vec!["SymbolValue|true"]);
}

#[test]
fn test_js_weakmap_registered_symbol_key_prohibited() {
    let src = r#"
const wm = new WeakMap();
const globalSym = Symbol.for("globalKey");
try {
    wm.set(globalSym, "Val");
} catch (e) {
    console.log("WeakMap Global Registered Symbol TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["WeakMap Global Registered Symbol TypeError"]
    );
}

#[test]
fn test_js_weakset_add_has_delete_flow() {
    let src = r#"
const ws = new WeakSet();
const obj = { name: "Alice" };
ws.add(obj);

console.log(ws.has(obj));
ws.delete(obj);
console.log(ws.has(obj));
"#;
    assert_eq!(run_js(src), vec!["true", "false"]);
}

#[test]
fn test_js_weakset_primitive_element_throws_typeerror() {
    let src = r#"
const ws = new WeakSet();
try {
    ws.add(42);
} catch (e) {
    console.log("WeakSet Primitive Element TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["WeakSet Primitive Element TypeError"]);
}

#[test]
fn test_js_weakset_symbol_element_support_es2023() {
    let src = r#"
const ws = new WeakSet();
const sym = Symbol("weakSetSym");
ws.add(sym);
console.log(ws.has(sym));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_weakmap_no_size_property() {
    let src = r#"
const wm = new WeakMap();
console.log(wm.size === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_weakset_no_size_property() {
    let src = r#"
const ws = new WeakSet();
console.log(ws.size === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_weakmap_no_iteration_methods() {
    let src = r#"
const wm = new WeakMap();
console.log((wm.keys === undefined) + "|" + (wm.values === undefined) + "|" + (wm.forEach === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_weakset_no_iteration_methods() {
    let src = r#"
const ws = new WeakSet();
console.log((ws.keys === undefined) + "|" + (ws.values === undefined) + "|" + (ws.forEach === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true|true"]);
}

#[test]
fn test_js_weakmap_chainable_set() {
    let src = r#"
const wm = new WeakMap();
const k1 = {}, k2 = {};
wm.set(k1, 1).set(k2, 2);
console.log(wm.get(k1) + "|" + wm.get(k2));
"#;
    assert_eq!(run_js(src), vec!["1|2"]);
}

#[test]
fn test_js_weakset_chainable_add() {
    let src = r#"
const ws = new WeakSet();
const o1 = {}, o2 = {};
ws.add(o1).add(o2);
console.log(ws.has(o1) + "|" + ws.has(o2));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_weakmap_constructor_initialization() {
    let src = r#"
const k1 = {}, k2 = {};
const wm = new WeakMap([[k1, "Val1"], [k2, "Val2"]]);
console.log(wm.get(k1) + "|" + wm.get(k2));
"#;
    assert_eq!(run_js(src), vec!["Val1|Val2"]);
}

#[test]
fn test_js_weakset_constructor_initialization() {
    let src = r#"
const o1 = {}, o2 = {};
const ws = new WeakSet([o1, o2]);
console.log(ws.has(o1) + "|" + ws.has(o2));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_weakmap_private_data_pattern_for_objects() {
    let src = r#"
const privateData = new WeakMap();
class Person {
    constructor(secret) {
        privateData.set(this, { secret });
    }
    getSecret() {
        return privateData.get(this).secret;
    }
}
const p = new Person("HiddenMessage");
console.log(p.getSecret());
"#;
    assert_eq!(run_js(src), vec!["HiddenMessage"]);
}

#[test]
fn test_js_weakset_brand_check_pattern() {
    let src = r#"
const brandStore = new WeakSet();
class CustomBrand {
    constructor() {
        brandStore.add(this);
    }
    static isInstance(obj) {
        return brandStore.has(obj);
    }
}
const b = new CustomBrand();
console.log(CustomBrand.isInstance(b) + "|" + CustomBrand.isInstance({}));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_weakmap_delete_returns_boolean() {
    let src = r#"
const wm = new WeakMap();
const k = {};
wm.set(k, 10);
console.log(wm.delete(k) + "|" + wm.delete(k));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_weakset_delete_returns_boolean() {
    let src = r#"
const ws = new WeakSet();
const o = {};
ws.add(o);
console.log(ws.delete(o) + "|" + ws.delete(o));
"#;
    assert_eq!(run_js(src), vec!["true|false"]);
}

#[test]
fn test_js_weakmap_function_as_key() {
    let src = r#"
const wm = new WeakMap();
function fnKey() {}
wm.set(fnKey, "FunctionMetaData");
console.log(wm.get(fnKey));
"#;
    assert_eq!(run_js(src), vec!["FunctionMetaData"]);
}
