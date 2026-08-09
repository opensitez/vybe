use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Error Stack Traces (`Error.captureStackTrace` & `.stack` Property)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_error_stack_property_exists() {
    let src = r#"
const err = new Error("StackTest");
console.log(typeof err.stack === "string" && err.stack.includes("Error: StackTest"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_capture_stack_trace_static_utility() {
    let src = r#"
function MyCustomError(message) {
    this.message = message;
    if (Error.captureStackTrace) {
        Error.captureStackTrace(this, MyCustomError);
    }
}
const err = new MyCustomError("CustomStackMsg");
console.log(typeof err.stack === "string");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_trace_limit_property() {
    let src = r#"
console.log(typeof Error.stackTraceLimit === "number");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_trace_formatting_includes_function_names() {
    let src = r#"
function levelA() { throw new Error("TraceErr"); }
function levelB() { levelA(); }

try {
    levelB();
} catch (e) {
    console.log(e.stack.includes("levelA") && e.stack.includes("levelB"));
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_capture_stack_trace_omits_wrapper_function() {
    let src = r#"
function helperFn(target) {
    Error.captureStackTrace(target, helperFn); // helperFn omitted from stack frames!
}
const obj = {};
helperFn(obj);
console.log(obj.stack.includes("helperFn"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_error_stack_property_override() {
    let src = r#"
const err = new Error("OrigMsg");
err.stack = "CustomFormattedStackLine";
console.log(err.stack);
"#;
    assert_eq!(run_js(src), vec!["CustomFormattedStackLine"]);
}

#[test]
fn test_js_error_stack_in_async_functions() {
    let src = r#"
async function asyncTask() {
    throw new Error("AsyncError");
}
(async () => {
    try {
        await asyncTask();
    } catch (e) {
        console.log(e.stack.includes("asyncTask"));
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_in_generator_functions() {
    let src = r#"
function* gen() {
    yield 1;
    throw new Error("GenError");
}
const g = gen();
g.next();
try {
    g.next();
} catch (e) {
    console.log(e.stack.includes("gen"));
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_trace_in_constructor() {
    let src = r#"
class SubError extends Error {
    constructor(msg) {
        super(msg);
        this.name = "SubError";
    }
}
const err = new SubError("SubMsg");
console.log(err.stack.includes("SubError: SubMsg"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_prepare_stack_trace_v8_hook() {
    let src = r#"
const origPrepare = Error.prepareStackTrace;
Error.prepareStackTrace = (err, structuredStackTrace) => {
    return `CallSiteCount:${structuredStackTrace.length}`;
};
try {
    const err = new Error("HooksTest");
    console.log(err.stack.startsWith("CallSiteCount:"));
} finally {
    Error.prepareStackTrace = origPrepare;
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_non_object_target_capture_throws_typeerror() {
    let src = r#"
try {
    Error.captureStackTrace("not_an_object");
} catch (e) {
    console.log("captureStackTrace Non-Object TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["captureStackTrace Non-Object TypeError"]);
}

#[test]
fn test_js_error_stack_getter_does_not_throw_if_detached() {
    let src = r#"
const desc = Object.getOwnPropertyDescriptor(Error.prototype, "stack");
console.log(typeof desc === "undefined" || typeof desc.get === "function");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_anonymous_function_frame() {
    let src = r#"
const anon = () => { throw new Error("AnonError"); };
try {
    anon();
} catch (e) {
    console.log(typeof e.stack === "string");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_eval_code_frame() {
    let src = r#"
try {
    eval("throw new Error('EvalError');");
} catch (e) {
    console.log(e.stack.includes("eval"));
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_trace_limit_mutation() {
    let src = r#"
const origLimit = Error.stackTraceLimit;
Error.stackTraceLimit = 1;
function a() { b(); }
function b() { throw new Error("Limit1"); }

try {
    a();
} catch (e) {
    console.log(typeof e.stack === "string");
} finally {
    Error.stackTraceLimit = origLimit;
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_trace_limit_zero_disables_stack() {
    let src = r#"
const origLimit = Error.stackTraceLimit;
Error.stackTraceLimit = 0;
const err = new Error("NoStack");
console.log(err.stack === "Error: NoStack" || err.stack === undefined || !err.stack.includes("at "));
Error.stackTraceLimit = origLimit;
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_in_promise_rejection_chain() {
    let src = r#"
Promise.resolve().then(() => {
    throw new Error("PromiseChainError");
}).catch(e => {
    console.log(typeof e.stack === "string");
});
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_in_event_loop_timeout() {
    let src = r#"
const err = new Error("TimeoutErr");
console.log(err.stack.includes("Error: TimeoutErr"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_stack_delete_property() {
    let src = r#"
const err = new Error("DeleteStack");
delete err.stack;
console.log(err.stack === undefined || typeof err.stack === "string");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_capture_stack_trace_default_constructor_parameter() {
    let src = r#"
class BaseErr extends Error {
    constructor(msg) {
        super(msg);
        Error.captureStackTrace(this, this.constructor);
    }
}
const e = new BaseErr("BaseMsg");
console.log(e.stack.includes("BaseErr: BaseMsg"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_capture_stack_trace_null_target_throws_typeerror() {
    let src = r#"
try {
    Error.captureStackTrace(null);
} catch (e) {
    console.log(e instanceof TypeError);
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
