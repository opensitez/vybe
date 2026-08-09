use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `Function.prototype.bind()`, Bound Functions & Currying
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_function_bind_explicit_this_context() {
    let src = r#"
function getName() {
    return this.name;
}
const user = { name: "Alice" };
const boundGetName = getName.bind(user);
console.log(boundGetName());
"#;
    assert_eq!(run_js(src), vec!["Alice"]);
}

#[test]
fn test_js_function_bind_partial_application_currying() {
    let src = r#"
function add(a, b, c) {
    return a + b + c;
}
const add5 = add.bind(null, 5);
const add5And10 = add5.bind(null, 10);
console.log(add5(2, 3) + "|" + add5And10(4));
"#;
    assert_eq!(run_js(src), vec!["10|19"]);
}

#[test]
fn test_js_bound_function_constructor_behavior_ignores_bound_this() {
    let src = r#"
function Point(x, y) {
    this.x = x;
    this.y = y;
}
const BoundPoint = Point.bind({ x: 99, y: 99 }, 10);
const p = new BoundPoint(20); // new operator ignores bound 'this' context, but retains prepended arguments!
console.log(`${p.x}:${p.y}`);
"#;
    assert_eq!(run_js(src), vec!["10:20"]);
}

#[test]
fn test_js_bound_function_has_no_prototype_property() {
    let src = r#"
function fn() {}
const bound = fn.bind(null);
console.log("prototype" in bound);
"#;
    assert_eq!(run_js(src), vec!["false"]); // Bound functions do not have a .prototype property!
}

#[test]
fn test_js_bound_function_name_prefix() {
    let src = r#"
function original() {}
const bound = original.bind(null);
console.log(bound.name);
"#;
    assert_eq!(run_js(src), vec!["bound original"]);
}

#[test]
fn test_js_bound_function_length_calculation() {
    let src = r#"
function sum(a, b, c, d) {}
const bound1 = sum.bind(null, 1);
const bound2 = sum.bind(null, 1, 2);
console.log(`${sum.length}:${bound1.length}:${bound2.length}`);
"#;
    assert_eq!(run_js(src), vec!["4:3:2"]);
}

#[test]
fn test_js_bound_function_length_clamped_to_zero() {
    let src = r#"
function sum(a, b) {}
const bound = sum.bind(null, 1, 2, 3);
console.log(bound.length);
"#;
    assert_eq!(run_js(src), vec!["0"]);
}

#[test]
fn test_js_bound_function_chaining_bind_does_not_override_this() {
    let src = r#"
function getThisName() {
    return this.name;
}
const obj1 = { name: "First" };
const obj2 = { name: "Second" };

const bound1 = getThisName.bind(obj1);
const bound2 = bound1.bind(obj2); // Re-binding does NOT change the initial 'this' binding!
console.log(bound2());
"#;
    assert_eq!(run_js(src), vec!["First"]);
}

#[test]
fn test_js_bound_function_call_and_apply_cannot_override_bound_this() {
    let src = r#"
function show() {
    return this.val;
}
const bound = show.bind({ val: "BoundVal" });
console.log(bound.call({ val: "CallVal" }) + "|" + bound.apply({ val: "ApplyVal" }));
"#;
    assert_eq!(run_js(src), vec!["BoundVal|BoundVal"]);
}

#[test]
fn test_js_bound_function_instanceof_checks_target_prototype() {
    let src = r#"
function TargetClass() {}
const BoundClass = TargetClass.bind(null);
const obj = new BoundClass();
console.log(obj instanceof TargetClass);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_bound_arrow_function_ignores_bound_this() {
    let src = r#"
const arrow = () => this;
const bound = arrow.bind({ a: 1 });
console.log(bound() === this); // Arrow function 'this' is lexically static, bind thisArg is ignored!
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_bound_arrow_function_prepends_arguments() {
    let src = r#"
const arrow = (a, b) => a + b;
const bound = arrow.bind(null, 10);
console.log(bound(5));
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_function_bind_primitive_this_boxed_in_non_strict() {
    let src = r#"
function checkThisType() {
    return typeof this;
}
const boundNum = checkThisType.bind(42);
console.log(boundNum()); // In non-strict mode, primitive this is boxed to Number object!
"#;
    assert_eq!(run_js(src), vec!["object"]);
}

#[test]
fn test_js_function_bind_primitive_this_unboxed_in_strict_mode() {
    let src = r#"
function checkThisType() {
    "use strict";
    return typeof this;
}
const boundNum = checkThisType.bind(42);
console.log(boundNum()); // In strict mode, primitive this remains primitive number!
"#;
    assert_eq!(run_js(src), vec!["number"]);
}

#[test]
fn test_js_function_bind_null_undefined_this_in_non_strict_becomes_globalthis() {
    let src = r#"
function getThisGlobal() {
    return this === globalThis;
}
const boundNull = getThisGlobal.bind(null);
console.log(boundNull());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_function_bind_null_this_in_strict_remains_null() {
    let src = r#"
function getThisNull() {
    "use strict";
    return this === null;
}
const boundNull = getThisNull.bind(null);
console.log(boundNull());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_bound_generator_function_returns_generator() {
    let src = r#"
function* gen(a, b) {
    yield a * this.factor;
    yield b * this.factor;
}
const boundGen = gen.bind({ factor: 10 }, 2);
const g = boundGen(3);
console.log([...g].join(","));
"#;
    assert_eq!(run_js(src), vec!["20,30"]);
}

#[test]
fn test_js_bound_async_function_returns_promise() {
    let src = r#"
async function asyncFn(a) {
    return a + this.val;
}
const boundAsync = asyncFn.bind({ val: 5 }, 10);
(async () => {
    console.log(await boundAsync());
})();
"#;
    assert_eq!(run_js(src), vec!["15"]);
}

#[test]
fn test_js_bound_function_anonymous_target_name() {
    let src = r#"
const bound = (function() {}).bind(null);
console.log(bound.name);
"#;
    assert_eq!(run_js(src), vec!["bound "]);
}

#[test]
fn test_js_bind_non_function_throws_typeerror() {
    let src = r#"
try {
    Function.prototype.bind.call("not_a_function");
} catch (e) {
    console.log("bind Non-Function TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["bind Non-Function TypeError"]);
}

#[test]
fn test_js_bind_getter_method_borrow() {
    let src = r#"
const obj = { get val() { return this._v; } };
const getter = Object.getOwnPropertyDescriptor(obj, "val").get.bind({ _v: 42 });
console.log(getter());
"#;
    assert_eq!(run_js(src), vec!["42"]);
}
