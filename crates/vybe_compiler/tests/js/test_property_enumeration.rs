/// Property enumeration, property order, Object.keys ordering rules,
/// for-in with prototype chain, string vs integer indices ordering.

use super::helpers::run_js;

// ── property ordering rules ───────────────────────────────────────────────────

#[test]
fn integer_indices_come_before_string_keys() {
    assert_eq!(run_js(r#"
const obj = {};
obj.b = 1;
obj[2] = 2;
obj.a = 3;
obj[1] = 4;
obj[0] = 5;
// Integer keys first (ascending), then string keys in insertion order
const keys = Object.keys(obj);
console.log(keys.join(","));
"#), vec!["0,1,2,b,a"]);
}

#[test]
fn symbol_keys_not_in_object_keys() {
    assert_eq!(run_js(r#"
const sym = Symbol("x");
const obj = { [sym]: 1, a: 2, b: 3 };
console.log(Object.keys(obj).join(","));
console.log(Object.getOwnPropertySymbols(obj).length);
"#), vec!["a,b", "1"]);
}

// ── for-in traversal ─────────────────────────────────────────────────────────

#[test]
fn for_in_traverses_prototype_chain() {
    assert_eq!(run_js(r#"
const proto = { inherited: "yes" };
const obj = Object.create(proto);
obj.own = "yes";
const found = {};
for (const k in obj) found[k] = true;
console.log(found.own);
console.log(found.inherited);
"#), vec!["true", "true"]);
}

#[test]
fn for_in_skips_non_enumerable_inherited() {
    assert_eq!(run_js(r#"
const obj = {};
// toString is non-enumerable on Object.prototype
const keys = [];
for (const k in obj) keys.push(k);
console.log(keys.includes("toString"));
"#), vec!["false"]);
}

// ── Object.keys vs getOwnPropertyNames ───────────────────────────────────────

#[test]
fn get_own_property_names_includes_non_enumerable() {
    assert_eq!(run_js(r#"
const obj = { a: 1 };
Object.defineProperty(obj, "b", { value: 2, enumerable: false });
const all = Object.getOwnPropertyNames(obj).sort();
const enumOnly = Object.keys(obj);
console.log(all.join(","));
console.log(enumOnly.join(","));
"#), vec!["a,b", "a"]);
}

// ── Reflect.ownKeys includes symbols ─────────────────────────────────────────

#[test]
fn reflect_ownkeys_includes_symbols_and_all_strings() {
    assert_eq!(run_js(r#"
const sym = Symbol("s");
const obj = {};
Object.defineProperty(obj, "hidden", { value: 1, enumerable: false });
obj.visible = 2;
obj[sym] = 3;
const all = Reflect.ownKeys(obj);
console.log(all.includes("hidden"));
console.log(all.includes("visible"));
console.log(all.some(k => typeof k === "symbol"));
"#), vec!["true", "true", "true"]);
}

// ── property existence checks ─────────────────────────────────────────────────

#[test]
fn in_vs_hasownproperty_for_inherited() {
    assert_eq!(run_js(r#"
const obj = { own: 1 };
console.log("own" in obj);
console.log("toString" in obj);   // inherited
console.log(obj.hasOwnProperty("own"));
console.log(obj.hasOwnProperty("toString"));
"#), vec!["true", "true", "true", "false"]);
}

// ── Object.entries order ──────────────────────────────────────────────────────

#[test]
fn object_entries_follows_key_order() {
    assert_eq!(run_js(r#"
const obj = { z: 3, a: 1, m: 2 };
const entries = Object.entries(obj);
// Insertion order for non-integer keys
console.log(entries.map(([k]) => k).join(","));
"#), vec!["z,a,m"]);
}

// ── property deletion order preservation ──────────────────────────────────────

#[test]
fn deletion_doesnt_reorder_remaining_keys() {
    assert_eq!(run_js(r#"
const obj = { a: 1, b: 2, c: 3, d: 4 };
delete obj.b;
console.log(Object.keys(obj).join(","));
"#), vec!["a,c,d"]);
}

// ── spreading and key order ───────────────────────────────────────────────────

#[test]
fn spread_preserves_key_insertion_order() {
    assert_eq!(run_js(r#"
const base = { x: 1, y: 2 };
const merged = { ...base, z: 3, x: 99 }; // x overridden but order retained
console.log(Object.keys(merged).join(","));
console.log(merged.x);
"#), vec!["x,y,z", "99"]);
}

// ── array index ordering ──────────────────────────────────────────────────────

#[test]
fn array_indices_sorted_as_integers() {
    assert_eq!(run_js(r#"
const arr = {};
arr[100] = "c";
arr[2] = "b";
arr[1] = "a";
arr.extra = "e";
console.log(Object.keys(arr).join(","));
"#), vec!["1,2,100,extra"]);
}

// ── JSON.stringify key ordering ───────────────────────────────────────────────

#[test]
fn json_stringify_preserves_insertion_order() {
    assert_eq!(run_js(r#"
const obj = { c: 3, a: 1, b: 2 };
const json = JSON.stringify(obj);
// JSON preserves insertion order
console.log(json);
"#), vec!["{\"c\":3,\"a\":1,\"b\":2}"]);
}

// ── property enumeration with proxy ──────────────────────────────────────────

#[test]
fn proxy_ownkeys_can_reorder() {
    assert_eq!(run_js(r#"
const target = { a: 1, b: 2, c: 3 };
const proxy = new Proxy(target, {
    ownKeys() { return ["c", "a", "b"]; }
});
const keys = Object.keys(proxy);
console.log(keys.join(","));
"#), vec!["c,a,b"]);
}
