use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Error Cause Chain Unwinding & Formatted Inspection
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_error_cause_chain_unwinding_traversal() {
    let src = r#"
function getCauseChain(err) {
    const chain = [];
    let current = err;
    while (current) {
        chain.push(current.message);
        current = current.cause;
    }
    return chain;
}
const dbErr = new Error("Database Connection Timeout");
const serviceErr = new Error("UserService Failed", { cause: dbErr });
const apiErr = new Error("HTTP 500 Internal Server Error", { cause: serviceErr });

console.log(getCauseChain(apiErr).join(" <- "));
"#;
    assert_eq!(
        run_js(src),
        vec!["HTTP 500 Internal Server Error <- UserService Failed <- Database Connection Timeout"]
    );
}

#[test]
fn test_js_error_cause_cycle_detection_in_unwinder() {
    let src = r#"
function getCauseChainSafe(err) {
    const visited = new Set();
    const chain = [];
    let current = err;
    while (current && !visited.has(current)) {
        visited.add(current);
        chain.push(current.message);
        current = current.cause;
    }
    return chain;
}
const e1 = new Error("E1");
const e2 = new Error("E2", { cause: e1 });
e1.cause = e2; // Cyclic cause chain!

console.log(getCauseChainSafe(e2).join(" -> "));
"#;
    assert_eq!(run_js(src), vec!["E2 -> E1"]);
}

#[test]
fn test_js_error_cause_chain_depth_limit() {
    let src = r#"
let err = new Error("Base");
for (let i = 1; i <= 5; i++) {
    err = new Error(`Layer ${i}`, { cause: err });
}
let depth = 0;
let cur = err;
while (cur) { depth++; cur = cur.cause; }
console.log(depth);
"#;
    assert_eq!(run_js(src), vec!["6"]);
}

#[test]
fn test_js_error_cause_heterogeneous_types_in_chain() {
    let src = r#"
const numCause = 404;
const strCause = "NOT_FOUND";
const err1 = new Error("HttpError", { cause: numCause });
const err2 = new Error("ApiError", { cause: strCause });

console.log(`${typeof err1.cause}:${err1.cause} | ${typeof err2.cause}:${err2.cause}`);
"#;
    assert_eq!(run_js(src), vec!["number:404 | string:NOT_FOUND"]);
}

#[test]
fn test_js_error_cause_object_descriptor_enumerable() {
    let src = r#"
const err = new Error("Msg", { cause: "Reason" });
const desc = Object.getOwnPropertyDescriptor(err, "cause");
console.log(desc.writable + "|" + desc.enumerable + "|" + desc.configurable);
"#;
    assert_eq!(run_js(src), vec!["true|false|true"]);
}

#[test]
fn test_js_error_cause_chain_aggregate_error_mix() {
    let src = r#"
const e1 = new Error("SubTaskA Failed");
const e2 = new Error("SubTaskB Failed");
const agg = new AggregateError([e1, e2], "Batch Failed");
const root = new Error("Job Failed", { cause: agg });

console.log(root.cause.errors.map(e => e.message).join(","));
"#;
    assert_eq!(run_js(src), vec!["SubTaskA Failed,SubTaskB Failed"]);
}

#[test]
fn test_js_error_cause_reasignment() {
    let src = r#"
const err = new Error("Original");
err.cause = "NewCauseReason";
console.log(err.cause);
"#;
    assert_eq!(run_js(src), vec!["NewCauseReason"]);
}

#[test]
fn test_js_error_cause_delete_property() {
    let src = r#"
const err = new Error("Msg", { cause: "Reason" });
delete err.cause;
console.log(Object.hasOwn(err, "cause"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_error_cause_chain_rethrow_wrapper_pattern() {
    let src = r#"
function runOperation() {
    try {
        JSON.parse("invalid_json");
    } catch (e) {
        throw new Error("Failed to parse config file", { cause: e });
    }
}
try {
    runOperation();
} catch (e) {
    console.log(e.message + "|isSyntaxError=" + (e.cause instanceof SyntaxError));
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Failed to parse config file|isSyntaxError=true"]
    );
}

#[test]
fn test_js_error_cause_chain_formatting_utility() {
    let src = r#"
function formatErrorChain(err) {
    let msg = err.name + ": " + err.message;
    if (err.cause) {
        msg += "\n  [caused by]: " + (err.cause instanceof Error ? formatErrorChain(err.cause) : err.cause);
    }
    return msg;
}
const inner = new TypeError("Invalid argument");
const outer = new Error("Action failed", { cause: inner });
console.log(formatErrorChain(outer));
"#;
    assert_eq!(
        run_js(src),
        vec!["Error: Action failed\n  [caused by]: TypeError: Invalid argument"]
    );
}

#[test]
fn test_js_error_cause_symbol_primitive() {
    let src = r#"
const sym = Symbol("err_code");
const err = new Error("SystemError", { cause: sym });
console.log(err.cause === sym);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_cause_null_vs_undefined() {
    let src = r#"
const eNull = new Error("Msg", { cause: null });
const eUndef = new Error("Msg", { cause: undefined });
console.log((eNull.cause === null) + "|" + (eUndef.cause === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_error_cause_with_getter_setter_on_subclass() {
    let src = r#"
class TrackedError extends Error {
    #cause;
    constructor(msg, cause) {
        super(msg);
        this.#cause = cause;
    }
    get cause() { return this.#cause; }
}
const err = new TrackedError("TrackedMsg", "InternalReason");
console.log(err.cause);
"#;
    assert_eq!(run_js(src), vec!["InternalReason"]);
}

#[test]
fn test_js_error_cause_json_stringify_custom_replacer() {
    let src = r#"
const cause = new Error("InnerMsg");
const err = new Error("OuterMsg", { cause });
const json = JSON.stringify(err, (key, value) => {
    if (value instanceof Error) {
        return { name: value.name, message: value.message, cause: value.cause };
    }
    return value;
});
console.log(json.includes("OuterMsg") && json.includes("InnerMsg"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_error_cause_promise_catch_chain() {
    let src = r#"
(async () => {
    try {
        await Promise.reject(new Error("AsyncLowLevel")).catch(e => {
            throw new Error("AsyncHighLevel", { cause: e });
        });
    } catch (e) {
        console.log(e.message + " <- " + e.cause.message);
    }
})();
"#;
    assert_eq!(run_js(src), vec!["AsyncHighLevel <- AsyncLowLevel"]);
}

#[test]
fn test_js_error_cause_non_object_options_ignored() {
    let src = r#"
const err = new Error("Msg", "not_an_object_options");
console.log(Object.hasOwn(err, "cause"));
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_error_cause_in_custom_subclass_super_forwarding() {
    let src = r#"
class AppError extends Error {
    constructor(message, options) {
        super(message, options);
    }
}
const err = new AppError("AppFailed", { cause: "DatabaseError" });
console.log(err.cause);
"#;
    assert_eq!(run_js(src), vec!["DatabaseError"]);
}

#[test]
fn test_js_error_cause_array_of_errors() {
    let src = r#"
const errors = [new Error("E1"), new Error("E2")];
const main = new Error("MultiFailure", { cause: errors });
console.log(main.cause.length);
"#;
    assert_eq!(run_js(src), vec!["2"]);
}

#[test]
fn test_js_error_cause_frozen_options_object() {
    let src = r#"
const opts = Object.freeze({ cause: "FrozenCause" });
const err = new Error("Msg", opts);
console.log(err.cause);
"#;
    assert_eq!(run_js(src), vec!["FrozenCause"]);
}

#[test]
fn test_js_error_cause_accessor_throwing_in_options() {
    let src = r#"
const opts = {
    get cause() { throw new Error("CauseGetterError"); }
};
try {
    new Error("Msg", opts);
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["CauseGetterError"]);
}

#[test]
fn test_js_error_cause_explicit_undefined_options() {
    let src = r#"
const err = new Error("Msg", { cause: undefined });
console.log(Object.hasOwn(err, "cause"));
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
