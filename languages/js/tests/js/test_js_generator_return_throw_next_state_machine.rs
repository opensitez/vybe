use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Generator Execution State Machine (`next()`, `return()`, `throw()`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_generator_basic_next_yield_values() {
    let src = r#"
function* gen() {
    yield 1;
    yield 2;
    return 3;
}
const g = gen();
console.log(`${JSON.stringify(g.next())}:${JSON.stringify(g.next())}:${JSON.stringify(g.next())}`);
"#;
    assert_eq!(
        run_js(src),
        vec![r#"{"value":1,"done":false}:{"value":2,"done":false}:{"value":3,"done":true}"#]
    );
}

#[test]
fn test_js_generator_passing_arguments_into_next() {
    let src = r#"
function* gen() {
    const a = yield "first";
    const b = yield a * 2;
    return b + 10;
}
const g = gen();
console.log(g.next().value); // "first"
console.log(g.next(5).value); // 5 * 2 = 10
console.log(g.next(20).value); // 20 + 10 = 30
"#;
    assert_eq!(run_js(src), vec!["first", "10", "30"]);
}

#[test]
fn test_js_generator_return_method_forces_completion() {
    let src = r#"
function* gen() {
    yield 10;
    yield 20;
}
const g = gen();
g.next();
const ret = g.return("ForcedReturn");
console.log(`${ret.value}:${ret.done}:${g.next().done}`);
"#;
    assert_eq!(run_js(src), vec!["ForcedReturn:true:true"]);
}

#[test]
fn test_js_generator_throw_method_handled_inside_generator() {
    let src = r#"
function* gen() {
    try {
        yield 1;
    } catch (e) {
        yield "HandledInGen: " + e.message;
    }
}
const g = gen();
g.next();
console.log(g.throw(new Error("ExternalError")).value);
"#;
    assert_eq!(run_js(src), vec!["HandledInGen: ExternalError"]);
}

#[test]
fn test_js_generator_throw_unhandled_propagates_exception() {
    let src = r#"
function* gen() {
    yield 1;
    yield 2;
}
const g = gen();
g.next();
try {
    g.throw(new Error("Unhandled"));
} catch (e) {
    console.log(e.message + "|done=" + g.next().done);
}
"#;
    assert_eq!(run_js(src), vec!["Unhandled|done=true"]);
}

#[test]
fn test_js_generator_finally_block_executes_on_return() {
    let src = r#"
let cleanedUp = false;
function* gen() {
    try {
        yield 1;
    } finally {
        cleanedUp = true;
    }
}
const g = gen();
g.next();
g.return();
console.log(cleanedUp);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_generator_finally_block_yields_on_return() {
    let src = r#"
function* gen() {
    try {
        yield 1;
    } finally {
        yield "FinallyYield";
    }
}
const g = gen();
g.next();
const ret = g.return("EarlyRet");
console.log(`${ret.value}:${ret.done}`); // Yield in finally intercepts return!
"#;
    assert_eq!(run_js(src), vec!["FinallyYield:false"]);
}

#[test]
fn test_js_generator_cannot_be_invoked_with_new_throws_typeerror() {
    let src = r#"
function* gen() {}
try {
    new gen();
} catch (e) {
    console.log("Generator Constructor TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Generator Constructor TypeError"]);
}

#[test]
fn test_js_generator_prototype_identity() {
    let src = r#"
function* gen() {}
const g = gen();
console.log(Object.getPrototypeOf(g) === gen.prototype);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_generator_tostringtag_is_generator() {
    let src = r#"
function* gen() {}
console.log(gen()[Symbol.toStringTag]);
"#;
    assert_eq!(run_js(src), vec!["Generator"]);
}

#[test]
fn test_js_generator_is_iterator_and_iterable() {
    let src = r#"
function* gen() { yield 1; }
const g = gen();
console.log(g[Symbol.iterator]() === g);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_generator_for_of_loop_consumption() {
    let src = r#"
function* gen() { yield "a"; yield "b"; yield "c"; }
const res = [];
for (const val of gen()) res.push(val);
console.log(res.join(","));
"#;
    assert_eq!(run_js(src), vec!["a,b,c"]);
}

#[test]
fn test_js_generator_array_spread_consumption() {
    let src = r#"
function* gen() { yield 10; yield 20; }
console.log([...gen()].join("-"));
"#;
    assert_eq!(run_js(src), vec!["10-20"]);
}

#[test]
fn test_js_generator_destructuring_consumption() {
    let src = r#"
function* gen() { yield 100; yield 200; }
const [a, b] = gen();
console.log(`${a}:${b}`);
"#;
    assert_eq!(run_js(src), vec!["100:200"]);
}

#[test]
fn test_js_generator_nested_try_catch_throw_handling() {
    let src = r#"
function* gen() {
    try {
        try {
            yield "inner";
        } catch (e) {
            yield "caughtInner: " + e.message;
            throw new Error("outerErr");
        }
    } catch (e) {
        yield "caughtOuter: " + e.message;
    }
}
const g = gen();
g.next();
g.throw(new Error("initial"));
console.log(g.next().value);
"#;
    assert_eq!(run_js(src), vec!["caughtOuter: outerErr"]);
}

#[test]
fn test_js_generator_exhausted_subsequent_next_calls() {
    let src = r#"
function* gen() { yield 1; }
const g = gen();
g.next(); // value: 1, done: false
g.next(); // value: undefined, done: true
const s3 = g.next(); // value: undefined, done: true
console.log(`${s3.value}:${s3.done}`);
"#;
    assert_eq!(run_js(src), vec!["undefined:true"]);
}

#[test]
fn test_js_generator_this_binding_in_method() {
    let src = r#"
const obj = {
    multiplier: 10,
    *gen() {
        yield 1 * this.multiplier;
        yield 2 * this.multiplier;
    }
};
console.log([...obj.gen()].join(","));
"#;
    assert_eq!(run_js(src), vec!["10,20"]);
}

#[test]
fn test_js_generator_expression_anonymous() {
    let src = r#"
const genFn = function*() { yield "AnonGen"; };
console.log(genFn().next().value);
"#;
    assert_eq!(run_js(src), vec!["AnonGen"]);
}

#[test]
fn test_js_generator_return_without_yield() {
    let src = r#"
function* gen() { return "DirectReturn"; }
console.log(JSON.stringify(gen().next()));
"#;
    assert_eq!(run_js(src), vec![r#"{"value":"DirectReturn","done":true}"#]);
}

#[test]
fn test_js_generator_throw_on_unstarted_generator() {
    let src = r#"
function* gen() { yield 1; }
const g = gen();
try {
    g.throw(new Error("UnstartedThrow"));
} catch (e) {
    console.log(e.message + "|done=" + g.next().done);
}
"#;
    assert_eq!(run_js(src), vec!["UnstartedThrow|done=true"]);
}

#[test]
fn test_js_generator_return_unstarted_generator() {
    let src = r#"
function* gen() { yield 1; }
const g = gen();
const r = g.return("early");
console.log(`${r.value}:${r.done}:${g.next().done}`);
"#;
    assert_eq!(run_js(src), vec!["early:true:true"]);
}

