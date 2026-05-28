/// Class decorators (ES2026 / Stage 3) — class, method, field, accessor decorators,
/// decorator metadata, stacking, decorator factories.

use super::helpers::run_js;

// ── class decorators ──────────────────────────────────────────────────────────

#[test]
fn class_decorator_basic() {
    assert_eq!(run_js(r#"
function sealed(target, ctx) {
    Object.seal(target.prototype);
    return target;
}

@sealed
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}

const p = new Point(1, 2);
console.log(p.x);
console.log(Object.isSealed(Point.prototype));
"#), vec!["1", "true"]);
}

#[test]
fn class_decorator_replaces_class() {
    assert_eq!(run_js(r#"
function withVersion(target, ctx) {
    target.version = "1.0";
    return target;
}

@withVersion
class App {}

console.log(App.version);
"#), vec!["1.0"]);
}

#[test]
fn class_decorator_factory() {
    assert_eq!(run_js(r#"
function tag(name) {
    return function(target, ctx) {
        target.tag = name;
    };
}

@tag("myClass")
class Foo {}

console.log(Foo.tag);
"#), vec!["myClass"]);
}

// ── method decorators ─────────────────────────────────────────────────────────

#[test]
fn method_decorator_wraps_method() {
    assert_eq!(run_js(r#"
function logged(fn, ctx) {
    const name = ctx.name;
    return function(...args) {
        console.log("call:" + name);
        return fn.apply(this, args);
    };
}

class Calculator {
    @logged
    add(a, b) { return a + b; }
}

const c = new Calculator();
console.log(c.add(2, 3));
"#), vec!["call:add", "5"]);
}

#[test]
fn method_decorator_memoize() {
    assert_eq!(run_js(r#"
function memoize(fn, ctx) {
    const cache = new Map();
    return function(x) {
        if (cache.has(x)) { console.log("cached"); return cache.get(x); }
        const result = fn.call(this, x);
        cache.set(x, result);
        return result;
    };
}

class Math2 {
    @memoize
    square(n) { return n * n; }
}

const m = new Math2();
console.log(m.square(4));
console.log(m.square(4));
"#), vec!["16", "cached", "16"]);
}

// ── accessor decorators ───────────────────────────────────────────────────────

#[test]
fn accessor_decorator_basic() {
    assert_eq!(run_js(r#"
function clamp(min, max) {
    return function(target, ctx) {
        const { get, set } = target;
        return {
            get() { return get.call(this); },
            set(v) { set.call(this, Math.min(max, Math.max(min, v))); }
        };
    };
}

class Slider {
    #value = 0;

    @clamp(0, 100)
    get value() { return this.#value; }
    set value(v) { this.#value = v; }
}

const s = new Slider();
s.value = 150;
console.log(s.value);
s.value = -10;
console.log(s.value);
"#), vec!["100", "0"]);
}

// ── field decorators ──────────────────────────────────────────────────────────

#[test]
fn field_decorator_initializer() {
    assert_eq!(run_js(r#"
function defaultValue(val) {
    return function(target, ctx) {
        return function() { return val; };
    };
}

class Config {
    @defaultValue(42)
    timeout;
}

const c = new Config();
console.log(c.timeout);
"#), vec!["42"]);
}

// ── stacking decorators ───────────────────────────────────────────────────────

#[test]
fn stacked_method_decorators_apply_inside_out() {
    assert_eq!(run_js(r#"
function addA(fn, ctx) { return function(...args) { return fn.apply(this, args) + "A"; }; }
function addB(fn, ctx) { return function(...args) { return fn.apply(this, args) + "B"; }; }

class Str {
    @addA
    @addB
    hello() { return "X"; }
}

// Decorators apply bottom-up: addB wraps hello first, then addA wraps the result
// hello() → "X", addB → "XB", addA → "XBA"
console.log(new Str().hello());
"#), vec!["XBA"]);
}

// ── decorator context object ──────────────────────────────────────────────────

#[test]
fn decorator_context_has_kind_and_name() {
    assert_eq!(run_js(r#"
const info = [];
function inspect(fn, ctx) {
    info.push(ctx.kind + ":" + ctx.name);
    return fn;
}

class MyClass {
    @inspect
    myMethod() {}
}

console.log(info[0]);
"#), vec!["method:myMethod"]);
}

#[test]
fn class_decorator_context_kind_is_class() {
    assert_eq!(run_js(r#"
let kind;
function capture(target, ctx) { kind = ctx.kind; }

@capture
class Foo {}

console.log(kind);
"#), vec!["class"]);
}

// ── decorator with metadata ───────────────────────────────────────────────────

#[test]
fn decorator_addInitializer_runs_after_class() {
    assert_eq!(run_js(r#"
const log = [];

function setup(fn, ctx) {
    ctx.addInitializer(function() {
        log.push("init:" + ctx.name);
    });
    return fn;
}

class Service {
    @setup
    start() {}
}

console.log(log.join(","));
"#), vec!["init:start"]);
}

// ── static method decorators ──────────────────────────────────────────────────

#[test]
fn static_method_decorator() {
    assert_eq!(run_js(r#"
function once(fn, ctx) {
    let called = false, result;
    return function(...args) {
        if (!called) { called = true; result = fn.apply(this, args); }
        return result;
    };
}

class Factory {
    static count = 0;

    @once
    static create() { return ++Factory.count; }
}

console.log(Factory.create());
console.log(Factory.create());
console.log(Factory.count);
"#), vec!["1", "1", "1"]);
}
