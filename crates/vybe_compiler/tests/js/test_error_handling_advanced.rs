/// Error handling edge cases and patterns

use super::helpers::run_js;

#[test]
fn custom_error_hierarchy() {
    assert_eq!(run_js(r#"
class AppError extends Error {
    constructor(message, code) {
        super(message);
        this.name = "AppError";
        this.code = code;
    }
}
class NetworkError extends AppError {
    constructor(message, statusCode) {
        super(message, "NETWORK");
        this.name = "NetworkError";
        this.statusCode = statusCode;
    }
}
const e = new NetworkError("Not Found", 404);
console.log(e instanceof Error);
console.log(e instanceof AppError);
console.log(e instanceof NetworkError);
console.log(e.code);
console.log(e.statusCode);
console.log(e.message);
"#), vec!["true", "true", "true", "NETWORK", "404", "Not Found"]);
}

#[test]
fn error_type_discrimination() {
    assert_eq!(run_js(r#"
function classify(fn) {
    try {
        fn();
    } catch(e) {
        if (e instanceof TypeError) return "type";
        if (e instanceof RangeError) return "range";
        if (e instanceof SyntaxError) return "syntax";
        return "other:" + e.constructor.name;
    }
}
console.log(classify(() => null.x));
console.log(classify(() => new Array(-1)));
console.log(classify(() => { throw new RangeError("oops"); }));
"#), vec!["type", "range", "range"]);
}

#[test]
fn try_catch_in_promise_chain() {
    assert_eq!(run_js(r#"
async function main() {
    const result = await Promise.resolve(1)
        .then(v => { throw new Error("fail"); })
        .catch(e => "caught: " + e.message)
        .then(v => v + " recovered");
    console.log(result);
}
main();
"#), vec!["caught: fail recovered"]);
}

#[test]
fn error_in_finally() {
    assert_eq!(run_js(r#"
function test() {
    try {
        throw new Error("original");
    } finally {
        return "from finally";
    }
}
console.log(test());
"#), vec!["from finally"]);
}

#[test]
fn aggregate_error_catching() {
    assert_eq!(run_js(r#"
async function main() {
    try {
        await Promise.any([
            Promise.reject(new Error("e1")),
            Promise.reject(new Error("e2")),
        ]);
    } catch(e) {
        console.log(e instanceof AggregateError);
        console.log(e.errors.length);
        console.log(e.errors[0].message);
    }
}
main();
"#), vec!["true", "2", "e1"]);
}

#[test]
fn error_propagation_through_callbacks() {
    assert_eq!(run_js(r#"
function safe(fn) {
    try { return { ok: true, value: fn() }; }
    catch(e) { return { ok: false, error: e.message }; }
}
const r1 = safe(() => JSON.parse('{"x":1}'));
const r2 = safe(() => JSON.parse("invalid"));
console.log(r1.ok);
console.log(r1.value.x);
console.log(r2.ok);
console.log(typeof r2.error);
"#), vec!["true", "1", "false", "string"]);
}

#[test]
fn stack_overflow_detection() {
    assert_eq!(run_js(r#"
function recurse(n) {
    try { return recurse(n + 1); }
    catch(e) { return n; }
}
const depth = recurse(0);
console.log(depth > 100);
console.log(typeof depth);
"#), vec!["true", "number"]);
}

#[test]
fn error_cause_chain() {
    assert_eq!(run_js(r#"
function level3() { throw new Error("level3 fail"); }
function level2() {
    try { level3(); }
    catch(e) { throw new Error("level2 fail", { cause: e }); }
}
function level1() {
    try { level2(); }
    catch(e) { throw new Error("level1 fail", { cause: e }); }
}
try { level1(); } catch(e) {
    console.log(e.message);
    console.log(e.cause.message);
    console.log(e.cause.cause.message);
}
"#), vec!["level1 fail", "level2 fail", "level3 fail"]);
}

#[test]
fn unhandled_rejection_pattern() {
    assert_eq!(run_js(r#"
async function main() {
    const results = await Promise.allSettled([
        Promise.resolve(1),
        Promise.reject(new Error("fail")),
        Promise.resolve(3),
    ]);
    const statuses = results.map(r => r.status);
    console.log(statuses.join(","));
    console.log(results[1].reason.message);
}
main();
"#), vec!["fulfilled,rejected,fulfilled", "fail"]);
}

#[test]
fn error_in_generator() {
    assert_eq!(run_js(r#"
function* gen() {
    try {
        yield 1;
        yield 2;
    } catch(e) {
        yield "caught: " + e.message;
    }
    yield 3;
}
const g = gen();
console.log(g.next().value);
console.log(g.throw(new Error("oops")).value);
console.log(g.next().value);
"#), vec!["1", "caught: oops", "3"]);
}

#[test]
fn optional_catch_binding() {
    assert_eq!(run_js(r#"
function safeParse(s) {
    try { return { ok: true, val: JSON.parse(s) }; }
    catch { return { ok: false }; }
}
console.log(safeParse('{"x":1}').ok);
console.log(safeParse("bad").ok);
"#), vec!["true", "false"]);
}
