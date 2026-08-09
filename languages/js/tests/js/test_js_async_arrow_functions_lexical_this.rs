use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Async Arrow Functions & Lexical `this` Binding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_async_arrow_function_basic_invocation() {
    let src = r#"
const add = async (a, b) => a + b;
add(10, 20).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_async_arrow_function_lexical_this_in_object_method() {
    let src = r#"
const obj = {
    multiplier: 3,
    calc(numbers) {
        // Async arrow preserves enclosing 'this'
        const fn = async (n) => n * this.multiplier;
        return Promise.all(numbers.map(fn));
    }
};
obj.calc([1, 2, 3]).then(results => console.log(results.join(",")));
"#;
    assert_eq!(run_js(src), vec!["3,6,9"]);
}

#[test]
fn test_js_async_arrow_function_lexical_arguments_binding() {
    let src = r#"
function outer(a, b) {
    const fn = async () => arguments[0] + arguments[1];
    return fn();
}
outer(100, 200).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["300"]);
}

#[test]
fn test_js_async_arrow_function_cannot_be_used_as_constructor() {
    let src = r#"
const AsyncArrow = async () => {};
try {
    new AsyncArrow();
} catch (e) {
    console.log("Async Arrow Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Async Arrow Constructor TypeError"]);
}

#[test]
fn test_js_async_arrow_function_no_prototype_property() {
    let src = r#"
const fn = async () => {};
console.log(fn.prototype === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_arrow_function_single_parameter_parentheses_optional() {
    let src = r#"
const doubleAsync = async x => x * 2;
doubleAsync(15).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_async_arrow_function_returning_object_literal() {
    let src = r#"
const makeUser = async (name, id) => ({ name, id });
makeUser("Alice", 1).then(user => console.log(`${user.name}:${user.id}`));
"#;
    assert_eq!(run_js(src), vec!["Alice:1"]);
}

#[test]
fn test_js_async_arrow_function_call_apply_bind_cannot_override_this() {
    let src = r#"
const obj1 = { name: "Obj1" };
const obj2 = { name: "Obj2" };

const fn = async function() {
    // Arrow function inside fn inherits fn's 'this'
    const arrow = async () => this.name;
    return arrow.call(obj2); // call attempt ignored for lexical this
};

fn.call(obj1).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Obj1"]);
}

#[test]
fn test_js_async_arrow_function_nested_in_class_constructor() {
    let src = r#"
class Task {
    constructor(id) {
        this.id = id;
        this.run = async () => `Task_${this.id}`;
    }
}
const t = new Task(99);
t.run().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["Task_99"]);
}

#[test]
fn test_js_async_arrow_function_in_array_higher_order_map() {
    let src = r#"
const ids = [1, 2, 3];
Promise.all(ids.map(async id => id * 10))
    .then(results => console.log(results.join(",")));
"#;
    assert_eq!(run_js(src), vec!["10,20,30"]);
}

#[test]
fn test_js_async_arrow_function_destructured_rest_parameters() {
    let src = r#"
const process = async ({ prefix }, ...items) => items.map(i => `${prefix}_${i}`).join(",");
process({ prefix: "TAG" }, "A", "B").then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["TAG_A,TAG_B"]);
}

#[test]
fn test_js_async_arrow_function_in_event_emitter_callback_simulation() {
    let src = r#"
class Handler {
    constructor() { this.count = 0; }
    register(dispatcher) {
        dispatcher(async () => {
            this.count += 10;
        });
    }
}
let cb;
const h = new Handler();
h.register(fn => { cb = fn; });

cb().then(() => console.log(h.count));
"#;
    assert_eq!(run_js(src), vec!["10"]);
}

#[test]
fn test_js_async_arrow_function_await_in_expression_body() {
    let src = r#"
const fetchVal = async () => await Promise.resolve("DirectExpression");
fetchVal().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["DirectExpression"]);
}

#[test]
fn test_js_async_arrow_function_try_catch_block_body() {
    let src = r#"
const safeDivide = async (a, b) => {
    try {
        if (b === 0) throw new Error("ZeroDivision");
        return a / b;
    } catch (e) {
        return e.message;
    }
};
safeDivide(10, 0).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["ZeroDivision"]);
}

#[test]
fn test_js_async_arrow_function_lexical_super_binding() {
    let src = r#"
class Base {
    async getName() { return "BaseName"; }
}
class Sub extends Base {
    async getName() {
        const getSuper = async () => await super.getName();
        return (await getSuper()) + "_Extended";
    }
}
new Sub().getName().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["BaseName_Extended"]);
}

#[test]
fn test_js_async_arrow_function_implicit_promise_return() {
    let src = r#"
const getPromise = async () => "WrappedImplicitly";
console.log(getPromise() instanceof Promise);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_async_arrow_function_immediately_invoked() {
    let src = r#"
(async () => "AsyncIIFE")().then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["AsyncIIFE"]);
}

#[test]
fn test_js_async_arrow_function_curried_higher_order() {
    let src = r#"
const multiplyBy = factor => async val => val * factor;
const timesFive = multiplyBy(5);
timesFive(6).then(res => console.log(res));
"#;
    assert_eq!(run_js(src), vec!["30"]);
}

#[test]
fn test_js_async_arrow_function_default_parameter_side_effects() {
    let src = r#"
let count = 0;
const getDefault = () => ++count;
const fn = async (val = getDefault()) => val * 10;

fn().then(r1 => {
    fn(100).then(r2 => {
        console.log(`${r1},${r2}|count=${count}`);
    });
});
"#;
    assert_eq!(run_js(src), vec!["10,1000|count=1"]);
}

#[test]
fn test_js_async_arrow_function_name_inference() {
    let src = r#"
const myAsyncFunc = async () => {};
console.log(myAsyncFunc.name);
"#;
    assert_eq!(run_js(src), vec!["myAsyncFunc"]);
}

#[test]
fn test_js_async_arrow_function_length_property() {
    let src = r#"
const fn = async (a, b = 1, c) => {};
console.log(fn.length);
"#;
    assert_eq!(run_js(src), vec!["1"]);
}
