/// Immutable data patterns — record-like updates, frozen structures, persistent data

use super::helpers::run_js;

#[test]
fn immutable_update_object() {
    assert_eq!(run_js(r#"
const user = Object.freeze({ name: "Alice", age: 30, active: true });
const updated = { ...user, age: 31 };
console.log(user.age);     // original unchanged
console.log(updated.age);
console.log(updated.name);
"#), vec!["30", "31", "Alice"]);
}

#[test]
fn immutable_update_nested() {
    assert_eq!(run_js(r#"
const state = Object.freeze({
    user: { name: "Bob", settings: { theme: "light" } },
    count: 0
});
const newState = {
    ...state,
    user: { ...state.user, settings: { ...state.user.settings, theme: "dark" } },
    count: state.count + 1
};
console.log(state.user.settings.theme);
console.log(newState.user.settings.theme);
console.log(newState.count);
"#), vec!["light", "dark", "1"]);
}

#[test]
fn immutable_array_push() {
    assert_eq!(run_js(r#"
const arr = Object.freeze([1, 2, 3]);
const newArr = [...arr, 4];
console.log(arr.length);
console.log(newArr.length);
console.log(newArr.join(","));
"#), vec!["3", "4", "1,2,3,4"]);
}

#[test]
fn immutable_array_remove() {
    assert_eq!(run_js(r#"
const arr = [1, 2, 3, 4, 5];
function removeAt(arr, index) {
    return [...arr.slice(0, index), ...arr.slice(index + 1)];
}
const result = removeAt(arr, 2);
console.log(arr.join(","));    // original unchanged
console.log(result.join(","));
"#), vec!["1,2,3,4,5", "1,2,4,5"]);
}

#[test]
fn immutable_array_update_element() {
    assert_eq!(run_js(r#"
const arr = [1, 2, 3, 4, 5];
function updateAt(arr, index, value) {
    return arr.map((x, i) => i === index ? value : x);
}
const result = updateAt(arr, 2, 99);
console.log(arr.join(","));
console.log(result.join(","));
"#), vec!["1,2,3,4,5", "1,2,99,4,5"]);
}

#[test]
fn immutable_map_update() {
    assert_eq!(run_js(r#"
const config = new Map([["theme", "light"], ["lang", "en"]]);
// Create new Map with update
const updated = new Map([...config, ["theme", "dark"]]);
console.log(config.get("theme"));
console.log(updated.get("theme"));
console.log(updated.get("lang"));
"#), vec!["light", "dark", "en"]);
}

#[test]
fn structural_sharing_via_object_create() {
    assert_eq!(run_js(r#"
const base = { common: "shared", own: "base" };
const variant = Object.create(base);
variant.own = "variant";
console.log(variant.common);  // from prototype
console.log(variant.own);     // own property
console.log(base.own);        // original unchanged
"#), vec!["shared", "variant", "base"]);
}

#[test]
fn record_update_pattern_with_class() {
    assert_eq!(run_js(r#"
class Record {
    constructor(data) { Object.assign(this, data); Object.freeze(this); }
    update(changes) { return new Record({ ...this, ...changes }); }
}
const r1 = new Record({ x: 1, y: 2, z: 3 });
const r2 = r1.update({ y: 99 });
console.log(r1.y);
console.log(r2.y);
console.log(r2.x);
"#), vec!["2", "99", "1"]);
}

#[test]
fn persistent_stack() {
    assert_eq!(run_js(r#"
class Stack {
    constructor(head = null, tail = null) { this.head = head; this.tail = tail; }
    push(val) { return new Stack(val, this); }
    pop() { return this.tail; }
    peek() { return this.head; }
    get isEmpty() { return this.head === null; }
}
const s0 = new Stack();
const s1 = s0.push(1);
const s2 = s1.push(2);
const s3 = s2.push(3);
console.log(s3.peek());
console.log(s2.peek());  // s2 still intact
console.log(s3.pop().peek());
"#), vec!["3", "2", "2"]);
}
