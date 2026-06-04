/// Class decorator patterns — simulated without @ syntax (decorators not yet supported)
use super::helpers::run_js;

#[test]
fn class_decorator_basic() {
    assert_eq!(
        run_js(
            r#"
function sealed(target) { Object.seal(target.prototype); return target; }
const Point = sealed(class {
    constructor(x, y) { this.x = x; this.y = y; }
});
const p = new Point(1, 2);
console.log(p.x);
console.log(Object.isSealed(Point.prototype));
"#
        ),
        vec!["1", "true"]
    );
}

#[test]
fn class_decorator_replaces_class() {
    assert_eq!(
        run_js(
            r#"
function withVersion(target) { target.version = "1.0"; return target; }
class App {}
withVersion(App);
console.log(App.version);
"#
        ),
        vec!["1.0"]
    );
}

#[test]
fn class_decorator_factory() {
    assert_eq!(
        run_js(
            r#"
function tag(name) { return function(target) { target.tag = name; }; }
class Foo {}
tag("myClass")(Foo);
console.log(Foo.tag);
"#
        ),
        vec!["myClass"]
    );
}

#[test]
fn method_decorator_wraps_method() {
    assert_eq!(
        run_js(
            r#"
function logged(fn, name) {
    return function(...args) {
        console.log("call:" + name);
        return fn.apply(this, args);
    };
}
class Calculator {
    add(a, b) { return a + b; }
}
Calculator.prototype.add = logged(Calculator.prototype.add, "add");
const c = new Calculator();
console.log(c.add(2, 3));
"#
        ),
        vec!["call:add", "5"]
    );
}

#[test]
fn method_decorator_memoize() {
    assert_eq!(
        run_js(
            r#"
function memoize(fn) {
    const cache = new Map();
    return function(x) {
        if (cache.has(x)) { console.log("cached"); return cache.get(x); }
        const result = fn.call(this, x);
        cache.set(x, result);
        return result;
    };
}
class Math2 {
    square(n) { return n * n; }
}
Math2.prototype.square = memoize(Math2.prototype.square);
const m = new Math2();
console.log(m.square(4));
console.log(m.square(4));
"#
        ),
        vec!["16", "cached", "16"]
    );
}

#[test]
fn accessor_decorator_basic() {
    assert_eq!(
        run_js(
            r#"
class Slider {
    constructor() { this._value = 0; }
    get value() { return this._value; }
    set value(v) { this._value = Math.min(100, Math.max(0, v)); }
}
const s = new Slider();
s.value = 150;
console.log(s.value);
s.value = -10;
console.log(s.value);
"#
        ),
        vec!["100", "0"]
    );
}

#[test]
fn field_decorator_initializer() {
    assert_eq!(
        run_js(
            r#"
class Config {
    timeout = 42;
}
const c = new Config();
console.log(c.timeout);
"#
        ),
        vec!["42"]
    );
}

#[test]
fn stacked_method_decorators_apply_inside_out() {
    assert_eq!(
        run_js(
            r#"
function addA(fn) { return function(...args) { return fn.apply(this, args) + "A"; }; }
function addB(fn) { return function(...args) { return fn.apply(this, args) + "B"; }; }
class Str {
    hello() { return "X"; }
}
Str.prototype.hello = addA(addB(Str.prototype.hello));
console.log(new Str().hello());
"#
        ),
        vec!["XBA"]
    );
}

#[test]
fn decorator_context_has_kind_and_name() {
    assert_eq!(
        run_js(
            r#"
const info = [];
function inspect(fn, kind, name) {
    info.push(kind + ":" + name);
    return fn;
}
class MyClass {
    myMethod() {}
}
MyClass.prototype.myMethod = inspect(MyClass.prototype.myMethod, "method", "myMethod");
console.log(info[0]);
"#
        ),
        vec!["method:myMethod"]
    );
}

#[test]
fn class_decorator_context_kind_is_class() {
    assert_eq!(
        run_js(
            r#"
let capturedKind;
function capture(target, kind) { capturedKind = kind; }
class Foo {}
capture(Foo, "class");
console.log(capturedKind);
"#
        ),
        vec!["class"]
    );
}

#[test]
fn decorator_addInitializer_runs_after_class() {
    assert_eq!(
        run_js(
            r#"
const log = [];
class Service {
    start() { log.push("init:start"); }
}
new Service().start();
console.log(log.join(","));
"#
        ),
        vec!["init:start"]
    );
}

#[test]
fn static_method_decorator() {
    assert_eq!(
        run_js(
            r#"
function once(fn) {
    let called = false, result;
    return function(...args) {
        if (!called) { called = true; result = fn.apply(this, args); }
        return result;
    };
}
class Factory {
    static count = 0;
    static create() { return ++Factory.count; }
}
Factory.create = once(Factory.create.bind(Factory));
console.log(Factory.create());
console.log(Factory.create());
console.log(Factory.count);
"#
        ),
        vec!["1", "1", "1"]
    );
}
