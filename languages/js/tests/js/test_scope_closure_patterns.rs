/// Scope and variable binding — closures, TDZ, let/const, function hoisting
use super::helpers::run_js;

#[test]
fn closure_counter_factory() {
    assert_eq!(
        run_js(
            r#"
function makeCounter(init = 0) {
    let count = init;
    return {
        increment: () => ++count,
        decrement: () => --count,
        reset: () => { count = init; },
        get: () => count,
    };
}
const c = makeCounter(10);
c.increment();
c.increment();
c.decrement();
console.log(c.get());
c.reset();
console.log(c.get());
"#
        ),
        vec!["11", "10"]
    );
}

#[test]
fn closure_captures_block_scope_let() {
    assert_eq!(
        run_js(
            r#"
const fns = [];
for (let i = 0; i < 5; i++) {
    fns.push(() => i);
}
console.log(fns.map(f => f()).join(","));
"#
        ),
        vec!["0,1,2,3,4"]
    );
}

#[test]
fn var_hoisting_in_function() {
    assert_eq!(
        run_js(
            r#"
function f() {
    console.log(x); // undefined (hoisted)
    var x = 42;
    console.log(x);
}
f();
"#
        ),
        vec!["undefined", "42"]
    );
}

#[test]
fn let_in_different_blocks_independent() {
    assert_eq!(
        run_js(
            r#"
let x = "outer";
{
    let x = "inner1";
    console.log(x);
}
{
    let x = "inner2";
    console.log(x);
}
console.log(x);
"#
        ),
        vec!["inner1", "inner2", "outer"]
    );
}

#[test]
fn closure_over_mutable_variable() {
    assert_eq!(
        run_js(
            r#"
function makeAdder(base) {
    return function(n) {
        base += n;
        return base;
    };
}
const add = makeAdder(10);
console.log(add(5));
console.log(add(3));
console.log(add(-2));
"#
        ),
        vec!["15", "18", "16"]
    );
}

#[test]
fn immediately_invoked_arrow_with_side_effects() {
    assert_eq!(
        run_js(
            r#"
const result = (() => {
    const a = 1, b = 2, c = 3;
    return { a, b, c, sum: a + b + c };
})();
console.log(result.sum);
console.log(result.a);
"#
        ),
        vec!["6", "1"]
    );
}

#[test]
fn function_expression_in_block() {
    assert_eq!(
        run_js(
            r#"
{
    const fn = function named() { return 42; };
    console.log(fn());
    console.log(fn.name);
}
"#
        ),
        vec!["42", "named"]
    );
}

#[test]
fn nested_closure_scope_chain() {
    assert_eq!(
        run_js(
            r#"
const a = 1;
function level1() {
    const b = 2;
    function level2() {
        const c = 3;
        function level3() {
            return a + b + c;
        }
        return level3();
    }
    return level2();
}
console.log(level1());
"#
        ),
        vec!["6"]
    );
}

#[test]
fn const_requires_initializer() {
    assert_eq!(
        run_js(
            r#"
let threw = false;
try {
    eval("const x;");
} catch {
    threw = true;
}
console.log(threw);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn closure_in_event_handler_pattern() {
    assert_eq!(
        run_js(
            r#"
function attachHandlers(items) {
    return items.map((item, index) => ({
        name: item,
        handler: () => `Clicked ${item} at index ${index}`
    }));
}
const handlers = attachHandlers(["a", "b", "c"]);
console.log(handlers[0].handler());
console.log(handlers[2].handler());
"#
        ),
        vec!["Clicked a at index 0", "Clicked c at index 2"]
    );
}

#[test]
fn generator_closure_state() {
    assert_eq!(
        run_js(
            r#"
function* statefulGen(start) {
    let n = start;
    while (true) {
        const reset = yield n;
        if (reset) n = start;
        else n++;
    }
}
const gen = statefulGen(0);
console.log(gen.next().value);
console.log(gen.next().value);
console.log(gen.next().value);
console.log(gen.next(true).value); // reset
console.log(gen.next().value);
"#
        ),
        vec!["0", "1", "2", "0", "1"]
    );
}

#[test]
fn closure_captures_default_parameter() {
    assert_eq!(
        run_js(
            r#"
function f(x = 10, g = () => x) {
    return g();
}
console.log(f());
console.log(f(20));
"#
        ),
        vec!["10", "20"]
    );
}

