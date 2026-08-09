/// Error handling patterns — custom error hierarchies, typed catch, retry, validation
use super::helpers::run_js;

#[test]
fn validation_error_with_field() {
    assert_eq!(
        run_js(
            r#"
class ValidationError extends Error {
    constructor(field, message) {
        super(message);
        this.name = "ValidationError";
        this.field = field;
    }
}
function validateAge(age) {
    if (typeof age !== "number") throw new ValidationError("age", "must be a number");
    if (age < 0) throw new ValidationError("age", "must be non-negative");
    return age;
}
try {
    validateAge(-1);
} catch (e) {
    console.log(e instanceof ValidationError);
    console.log(e.field);
    console.log(e.message);
}
"#
        ),
        vec!["true", "age", "must be non-negative"]
    );
}

#[test]
fn retry_pattern() {
    assert_eq!(
        run_js(
            r#"
async function withRetry(fn, maxAttempts) {
    let lastError;
    for (let attempt = 1; attempt <= maxAttempts; attempt++) {
        try {
            return await fn(attempt);
        } catch (e) {
            lastError = e;
        }
    }
    throw lastError;
}
let callCount = 0;
async function main() {
    const result = await withRetry(async (attempt) => {
        callCount++;
        if (attempt < 3) throw new Error("not yet");
        return "success";
    }, 3);
    console.log(result);
    console.log(callCount);
}
main();
"#
        ),
        vec!["success", "3"]
    );
}

#[test]
fn error_boundary_pattern() {
    assert_eq!(
        run_js(
            r#"
function safe(fn) {
    return function(...args) {
        try {
            return { value: fn(...args), error: null };
        } catch (e) {
            return { value: null, error: e };
        }
    };
}
const safeDivide = safe((a, b) => {
    if (b === 0) throw new Error("division by zero");
    return a / b;
});
const r1 = safeDivide(10, 2);
const r2 = safeDivide(10, 0);
console.log(r1.value);
console.log(r1.error);
console.log(r2.value);
console.log(r2.error.message);
"#
        ),
        vec!["5", "null", "null", "division by zero"]
    );
}

#[test]
fn error_chain_via_cause() {
    assert_eq!(
        run_js(
            r#"
function fetchData() {
    throw new TypeError("network error");
}
function loadUser(id) {
    try {
        return fetchData(id);
    } catch (e) {
        throw new Error("Failed to load user " + id, { cause: e });
    }
}
try {
    loadUser(42);
} catch (e) {
    console.log(e.message);
    console.log(e.cause instanceof TypeError);
    console.log(e.cause.message);
}
"#
        ),
        vec!["Failed to load user 42", "true", "network error"]
    );
}

#[test]
fn multiple_catch_types() {
    assert_eq!(
        run_js(
            r#"
class NetworkError extends Error { constructor(m) { super(m); this.name = "NetworkError"; } }
class TimeoutError extends NetworkError { constructor(m) { super(m); this.name = "TimeoutError"; } }
function handle(err) {
    if (err instanceof TimeoutError) return "timeout";
    if (err instanceof NetworkError) return "network";
    if (err instanceof Error) return "error";
    return "unknown";
}
console.log(handle(new TimeoutError("slow")));
console.log(handle(new NetworkError("down")));
console.log(handle(new Error("oops")));
"#
        ),
        vec!["timeout", "network", "error"]
    );
}

#[test]
fn finally_cleanup_resource() {
    assert_eq!(
        run_js(
            r#"
const resources = [];
function openResource(id) {
    resources.push("open:" + id);
    return { id, close() { resources.push("close:" + id); } };
}
function process(id) {
    const res = openResource(id);
    try {
        if (id === 2) throw new Error("bad resource");
        return "ok";
    } finally {
        res.close();
    }
}
try { process(1); } catch {}
try { process(2); } catch {}
console.log(resources.join(","));
"#
        ),
        vec!["open:1,close:1,open:2,close:2"]
    );
}

#[test]
fn unhandled_rejection_can_be_caught() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    const results = await Promise.allSettled([
        Promise.reject(new Error("r1")),
        Promise.resolve("ok"),
        Promise.reject(new Error("r3")),
    ]);
    const errors = results.filter(r => r.status === "rejected").map(r => r.reason.message);
    console.log(errors.join(","));
}
main();
"#
        ),
        vec!["r1,r3"]
    );
}

#[test]
fn async_then_throw_handled_in_catch() {
    assert_eq!(
        run_js(
            r#"
async function main() {
    try {
        await Promise.resolve().then(() => {
            throw new Error("then_err");
        });
    } catch (e) {
        console.log(e.message);
    }
}
main();
"#
        ),
        vec!["then_err"]
    );
}
