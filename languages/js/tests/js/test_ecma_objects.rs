use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Objects — literals, methods, modern features
// ═══════════════════════════════════════════════════════════

#[test]
fn object_literal() {
    let out = run_js(
        r#"
const obj = { x: 1, y: 2, z: 3 };
console.log(obj.x + obj.y + obj.z);
"#,
    );
    assert_eq!(out, vec!["6"]);
}

#[test]
fn object_shorthand() {
    let out = run_js(
        r#"
const x = 10, y = 20;
const obj = { x, y };
console.log(obj.x);
console.log(obj.y);
"#,
    );
    assert_eq!(out, vec!["10", "20"]);
}

#[test]
fn object_computed_key() {
    let out = run_js(
        r#"
const key = "name";
const obj = { [key]: "Alice" };
console.log(obj.name);
"#,
    );
    assert_eq!(out, vec!["Alice"]);
}

#[test]
fn object_method_shorthand() {
    let out = run_js(
        r#"
const obj = {
    greet(name) { return "Hello " + name; },
    farewell(name) { return "Bye " + name; }
};
console.log(obj.greet("World"));
console.log(obj.farewell("World"));
"#,
    );
    assert_eq!(out, vec!["Hello World", "Bye World"]);
}

#[test]
fn object_nested() {
    let out = run_js(
        r#"
const obj = {
    a: { b: { c: 42 } }
};
console.log(obj.a.b.c);
"#,
    );
    assert_eq!(out, vec!["42"]);
}

#[test]
fn object_dynamic_access() {
    let out = run_js(
        r#"
const obj = { foo: 1, bar: 2 };
const key = "foo";
console.log(obj[key]);
obj["baz"] = 3;
console.log(obj.baz);
"#,
    );
    assert_eq!(out, vec!["1", "3"]);
}

#[test]
fn object_keys() {
    let out = run_js(
        r#"
const obj = { a: 1, b: 2, c: 3 };
const keys = Object.keys(obj);
console.log(keys.length);
"#,
    );
    assert_eq!(out, vec!["3"]);
}

#[test]
fn object_values() {
    let out = run_js(
        r#"
const obj = { a: 10, b: 20, c: 30 };
const vals = Object.values(obj);
let sum = 0;
for (const v of vals) sum += v;
console.log(sum);
"#,
    );
    assert_eq!(out, vec!["60"]);
}

#[test]
fn object_entries() {
    let out = run_js(
        r#"
const obj = { x: 1, y: 2 };
const entries = Object.entries(obj);
console.log(entries.length);
"#,
    );
    assert_eq!(out, vec!["2"]);
}

#[test]
fn object_assign() {
    let out = run_js(
        r#"
const target = { a: 1 };
const source = { b: 2, c: 3 };
Object.assign(target, source);
console.log(target.a);
console.log(target.b);
console.log(target.c);
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn object_spread() {
    let out = run_js(
        r#"
const a = { x: 1 };
const b = { y: 2 };
const merged = { ...a, ...b, z: 3 };
console.log(merged.x);
console.log(merged.y);
console.log(merged.z);
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn object_spread_override() {
    let out = run_js(
        r#"
const defaults = { color: "red", size: 10 };
const custom = { ...defaults, color: "blue" };
console.log(custom.color);
console.log(custom.size);
"#,
    );
    assert_eq!(out, vec!["blue", "10"]);
}

#[test]
fn hasownproperty() {
    let out = run_js(
        r#"
const obj = { a: 1 };
console.log(obj.hasOwnProperty("a"));
console.log(obj.hasOwnProperty("b"));
"#,
    );
    assert_eq!(out, vec!["true", "false"]);
}

#[test]
fn object_from_entries() {
    let out = run_js(
        r#"
const entries = [["a", 1], ["b", 2], ["c", 3]];
const obj = Object.fromEntries(entries);
console.log(obj.a);
console.log(obj.b);
console.log(obj.c);
"#,
    );
    assert_eq!(out, vec!["1", "2", "3"]);
}

#[test]
fn object_getter_literal() {
    let out = run_js(
        r#"
const obj = {
    _name: "test",
    get name() { return this._name.toUpperCase(); }
};
console.log(obj.name);
"#,
    );
    assert_eq!(out, vec!["TEST"]);
}

#[test]
fn object_setter_literal() {
    let out = run_js(
        r#"
const obj = {
    _val: 0,
    get val() { return this._val; },
    set val(v) { this._val = v * 2; }
};
obj.val = 5;
console.log(obj.val);
"#,
    );
    assert_eq!(out, vec!["10"]);
}

#[test]
fn object_pass_by_reference() {
    let out = run_js(
        r#"
function modify(obj) {
    obj.x = 99;
}
const o = { x: 1 };
modify(o);
console.log(o.x);
"#,
    );
    assert_eq!(out, vec!["99"]);
}

#[test]
fn object_destructure_assign() {
    let out = run_js(
        r#"
const { a, b, ...rest } = { a: 1, b: 2, c: 3, d: 4 };
console.log(a);
console.log(b);
"#,
    );
    assert_eq!(out, vec!["1", "2"]);
}

#[test]
fn object_spread_order_last_wins() {
    let out = run_js(
        r#"
const merged = { a: 1, ...{ a: 2, b: 3 }, a: 4 };
console.log(merged.a);
console.log(merged.b);
"#,
    );
    assert_eq!(out, vec!["4", "3"]);
}

#[test]
fn object_keys_ignore_symbol_properties() {
    let out = run_js(
        r#"
const id = Symbol("id");
const obj = { a: 1, [id]: 2 };
console.log(Object.keys(obj).join(","));
console.log(obj[id]);
"#,
    );
    assert_eq!(out, vec!["a", "2"]);
}

#[test]
fn object_entries_preserve_insertion_order() {
    let out = run_js(
        r#"
const obj = {};
obj.first = 1;
obj.second = 2;
obj.third = 3;
console.log(Object.entries(obj).map(([k, v]) => k + ":" + v).join(","));
"#,
    );
    assert_eq!(out, vec!["first:1,second:2,third:3"]);
}

#[test]
fn object_assign_returns_target_reference() {
    let out = run_js(
        r#"
const target = { a: 1 };
const result = Object.assign(target, { b: 2 });
console.log(result === target);
console.log(target.b);
"#,
    );
    assert_eq!(out, vec!["true", "2"]);
}

#[test]
fn object_from_entries_overwrites_duplicate_keys() {
    let out = run_js(
        r#"
const obj = Object.fromEntries([["a", 1], ["a", 2], ["b", 3]]);
console.log(obj.a);
console.log(obj.b);
"#,
    );
    assert_eq!(out, vec!["2", "3"]);
}

#[test]
fn object_literal_numeric_keys_access_as_strings() {
    let out = run_js(
        r#"
const obj = { 1: "one", 2: "two" };
console.log(obj[1]);
console.log(obj["2"]);
"#,
    );
    assert_eq!(out, vec!["one", "two"]);
}

#[test]
fn object_dynamic_property_delete_and_readd() {
    let out = run_js(
        r#"
const obj = { a: 1, b: 2 };
delete obj["a"];
obj["a"] = 3;
console.log(obj.a);
console.log(Object.keys(obj).join(","));
"#,
    );
    assert_eq!(out, vec!["3", "b,a"]);
}

#[test]
fn hasownproperty_distinguishes_inherited_members() {
    let out = run_js(
        r#"
const proto = { inherited: true };
const obj = Object.create(proto);
obj.own = true;
console.log(obj.hasOwnProperty("own"));
console.log(obj.hasOwnProperty("inherited"));
console.log("inherited" in obj);
"#,
    );
    assert_eq!(out, vec!["true", "false", "true"]);
}

#[test]
fn object_values_follow_object_keys_order() {
    let out = run_js(
        r#"
const obj = { a: 10, b: 20, c: 30 };
console.log(Object.values(obj).join(","));
"#,
    );
    assert_eq!(out, vec!["10,20,30"]);
}
