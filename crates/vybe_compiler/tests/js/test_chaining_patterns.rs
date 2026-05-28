/// Chaining patterns — method chaining, fluent interface, builder pattern

use super::helpers::run_js;

#[test]
fn method_chaining_fluent_interface() {
    assert_eq!(run_js(r#"
class StringBuilder {
    #parts = [];
    append(str) { this.#parts.push(str); return this; }
    prepend(str) { this.#parts.unshift(str); return this; }
    join(sep = "") { return this.#parts.join(sep); }
    toString() { return this.join(); }
}
const result = new StringBuilder()
    .append("World")
    .prepend("Hello ")
    .append("!")
    .toString();
console.log(result);
"#), vec!["Hello World!"]);
}

#[test]
fn lodash_chain_style() {
    assert_eq!(run_js(r#"
// Simplified lodash-like chain
class Chain {
    constructor(val) { this._val = val; }
    map(fn) { return new Chain(this._val.map(fn)); }
    filter(fn) { return new Chain(this._val.filter(fn)); }
    reduce(fn, init) { return this._val.reduce(fn, init); }
    value() { return this._val; }
}
const result = new Chain([1, 2, 3, 4, 5, 6])
    .filter(x => x % 2 === 0)
    .map(x => x * x)
    .reduce((a, b) => a + b, 0);
console.log(result); // 4+16+36 = 56
"#), vec!["56"]);
}

#[test]
fn promise_style_chaining() {
    assert_eq!(run_js(r#"
class Computation {
    constructor(val) { this._val = val; }
    map(fn) { return new Computation(fn(this._val)); }
    flatMap(fn) { return fn(this._val); }
    getOrElse(def) { return this._val ?? def; }
}
const result = new Computation(5)
    .map(x => x * 2)
    .map(x => x + 1)
    .flatMap(x => new Computation(x.toString()))
    .getOrElse("default");
console.log(result);
"#), vec!["11"]);
}

#[test]
fn sql_builder_chaining() {
    assert_eq!(run_js(r#"
class SQL {
    #parts = [];
    select(...cols) { this.#parts.push(`SELECT ${cols.join(", ")}`); return this; }
    from(table) { this.#parts.push(`FROM ${table}`); return this; }
    where(cond) { this.#parts.push(`WHERE ${cond}`); return this; }
    orderBy(col) { this.#parts.push(`ORDER BY ${col}`); return this; }
    limit(n) { this.#parts.push(`LIMIT ${n}`); return this; }
    build() { return this.#parts.join(" "); }
}
const query = new SQL()
    .select("id", "name")
    .from("users")
    .where("active = 1")
    .orderBy("name")
    .limit(10)
    .build();
console.log(query);
"#), vec!["SELECT id, name FROM users WHERE active = 1 ORDER BY name LIMIT 10"]);
}

#[test]
fn animation_chain() {
    assert_eq!(run_js(r#"
class Animation {
    #steps = [];
    moveTo(x, y) { this.#steps.push(`move(${x},${y})`); return this; }
    scaleTo(s) { this.#steps.push(`scale(${s})`); return this; }
    rotateTo(deg) { this.#steps.push(`rotate(${deg})`); return this; }
    play() { return this.#steps.join(" -> "); }
}
const anim = new Animation()
    .moveTo(100, 200)
    .scaleTo(2)
    .rotateTo(90)
    .moveTo(0, 0);
console.log(anim.play());
"#), vec!["move(100,200) -> scale(2) -> rotate(90) -> move(0,0)"]);
}

#[test]
fn array_method_chaining() {
    assert_eq!(run_js(r#"
const result = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]
    .filter(n => n % 2 === 0)
    .map(n => n * n)
    .reduce((sum, n) => sum + n, 0);
console.log(result); // 4+16+36+64+100 = 220
"#), vec!["220"]);
}

#[test]
fn string_method_chaining() {
    assert_eq!(run_js(r#"
const result = "  Hello, World!  "
    .trim()
    .toLowerCase()
    .replace(",", "")
    .split(" ")
    .join("-");
console.log(result);
"#), vec!["hello-world!"]);
}

#[test]
fn functional_chain_point_free() {
    assert_eq!(run_js(r#"
const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const process = pipe(
    arr => arr.filter(x => x > 2),
    arr => arr.map(x => x * 2),
    arr => arr.reduce((a, b) => a + b, 0)
);
console.log(process([1, 2, 3, 4, 5])); // [3,4,5] -> [6,8,10] -> 24
"#), vec!["24"]);
}
