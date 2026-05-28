/// Object.create, Object.assign, prototype-based OOP patterns

use super::helpers::run_js;

#[test]
fn object_create_null_safe_map() {
    assert_eq!(run_js(r#"
const map = {};
map.key = "value";
map.other = "data";
console.log(map.key);
console.log(Object.keys(map).length);
// User-added keys are own; inherited ones like toString are not own
console.log(Object.prototype.hasOwnProperty.call(map, "toString"));
"#), vec!["value", "2", "false"]);
}

#[test]
fn prototype_based_inheritance() {
    assert_eq!(run_js(r#"
const shape = {
    area() { return 0; },
    perimeter() { return 0; },
    describe() { return `${this.constructor.name}: area=${this.area()}`; }
};
const circle = Object.create(shape);
circle.constructor = { name: "Circle" };
circle.init = function(r) { this.r = r; return this; };
circle.area = function() { return Math.PI * this.r * this.r; };
const c = Object.create(circle).init(3);
console.log(c.area().toFixed(4));
"#), vec!["28.2743"]);
}

#[test]
fn assign_deep_merge_pattern() {
    assert_eq!(run_js(r#"
// Note: Object.assign does shallow copy
const target = { a: { x: 1, y: 2 }, b: 10 };
const source = { a: { z: 3 }, b: 20 };
const merged = Object.assign({}, target, source);
// a is overwritten (not deep merged)
console.log(merged.b);
console.log(merged.a.z); // has z from source
console.log(merged.a.x); // undefined — source.a replaced target.a
"#), vec!["20", "3", "undefined"]);
}

#[test]
fn object_create_with_accessor_in_props() {
    assert_eq!(run_js(r#"
const obj = Object.create({}, {
    x: {
        get() { return this._x ?? 0; },
        set(v) { this._x = v; },
        configurable: true,
        enumerable: true
    }
});
obj.x = 42;
console.log(obj.x);
"#), vec!["42"]);
}

#[test]
fn mixin_via_object_assign_to_prototype() {
    assert_eq!(run_js(r#"
const Serializable = {
    toJSON() { return JSON.stringify(this); },
    fromJSON(str) { return Object.assign(Object.create(this), JSON.parse(str)); }
};
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
Object.assign(Point.prototype, Serializable);
const p = new Point(3, 4);
const json = JSON.stringify(p);
const p2 = JSON.parse(json);
console.log(p2.x);
console.log(p2.y);
"#), vec!["3", "4"]);
}

#[test]
fn object_keys_values_entries_iteration() {
    assert_eq!(run_js(r#"
const config = { host: "localhost", port: 8080, debug: true };
const pairs = Object.entries(config)
    .map(([k, v]) => k + "=" + v)
    .join(",");
console.log(pairs);
"#), vec!["host=localhost,port=8080,debug=true"]);
}

#[test]
fn object_assign_clone_then_mutate() {
    assert_eq!(run_js(r#"
const original = { name: "Alice", score: 10 };
const copy = Object.assign({}, original);
copy.score += 5;
console.log(original.score);
console.log(copy.score);
"#), vec!["10", "15"]);
}

#[test]
fn null_prototype_used_as_pure_map() {
    assert_eq!(run_js(r#"
const counts = Object.create(null);
const words = ["apple", "banana", "apple", "cherry", "banana", "apple"];
for (const w of words) counts[w] = (counts[w] ?? 0) + 1;
console.log(counts.apple);
console.log(counts.banana);
console.log(counts.cherry);
"#), vec!["3", "2", "1"]);
}
