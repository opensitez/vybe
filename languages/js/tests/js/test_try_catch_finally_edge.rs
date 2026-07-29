/// try/catch/finally edge cases — return in finally overrides, nested try, re-throw patterns
use super::helpers::run_js;

#[test]
fn finally_runs_after_return() {
    assert_eq!(
        run_js(
            r#"
function f() {
    try {
        return "try";
    } finally {
        console.log("finally");
    }
}
console.log(f());
"#
        ),
        vec!["finally", "try"]
    );
}

#[test]
fn finally_return_overrides_try_return() {
    assert_eq!(
        run_js(
            r#"
function f() {
    try {
        return "try";
    } finally {
        return "finally"; // overrides!
    }
}
console.log(f());
"#
        ),
        vec!["finally"]
    );
}

#[test]
fn finally_runs_after_throw() {
    assert_eq!(
        run_js(
            r#"
let ran = false;
function f() {
    try {
        throw new Error("boom");
    } finally {
        ran = true;
    }
}
try { f(); } catch {}
console.log(ran);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn catch_receives_thrown_value() {
    assert_eq!(
        run_js(
            r#"
try {
    throw { code: 42, msg: "custom" };
} catch (e) {
    console.log(e.code);
    console.log(e.msg);
}
"#
        ),
        vec!["42", "custom"]
    );
}

#[test]
fn catch_binding_optional() {
    assert_eq!(
        run_js(
            r#"
try {
    throw new Error("ignored");
} catch {
    console.log("caught without binding");
}
"#
        ),
        vec!["caught without binding"]
    );
}

#[test]
fn rethrow_propagates_original() {
    assert_eq!(
        run_js(
            r#"
function inner() { throw new TypeError("inner"); }
function outer() {
    try { inner(); }
    catch (e) {
        if (!(e instanceof TypeError)) throw e;
        console.log("handled: " + e.message);
    }
}
outer();
"#
        ),
        vec!["handled: inner"]
    );
}

#[test]
fn nested_try_inner_catches_first() {
    assert_eq!(
        run_js(
            r#"
let order = [];
try {
    try {
        throw new Error("boom");
    } catch (e) {
        order.push("inner catch");
        throw e; // re-throw
    }
} catch (e) {
    order.push("outer catch");
}
console.log(order.join(","));
"#
        ),
        vec!["inner catch,outer catch"]
    );
}

#[test]
fn finally_executes_even_without_error() {
    assert_eq!(
        run_js(
            r#"
const log = [];
try {
    log.push("try");
} catch {
    log.push("catch");
} finally {
    log.push("finally");
}
console.log(log.join(","));
"#
        ),
        vec!["try,finally"]
    );
}

#[test]
fn error_in_catch_propagates_outside() {
    assert_eq!(
        run_js(
            r#"
let caught = false;
try {
    try {
        throw new Error("first");
    } catch {
        throw new Error("second"); // error in catch
    }
} catch (e) {
    caught = true;
    console.log(e.message);
}
console.log(caught);
"#
        ),
        vec!["second", "true"]
    );
}

#[test]
fn finally_does_not_suppress_throw() {
    assert_eq!(
        run_js(
            r#"
let caught = null;
try {
    try {
        throw new Error("original");
    } finally {
        // finally without catch — error still propagates
        console.log("finally runs");
    }
} catch (e) {
    caught = e.message;
}
console.log(caught);
"#
        ),
        vec!["finally runs", "original"]
    );
}

#[test]
fn try_catch_in_loop() {
    assert_eq!(
        run_js(
            r#"
const results = [];
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) throw new Error("skip");
        results.push(i);
    } catch {
        results.push("err");
    }
}
console.log(results.join(","));
"#
        ),
        vec!["0,err,2"]
    );
}

#[test]
fn throw_string_caught() {
    assert_eq!(
        run_js(
            r#"
try {
    throw "string error";
} catch (e) {
    console.log(typeof e);
    console.log(e);
}
"#
        ),
        vec!["string", "string error"]
    );
}

#[test]
fn throw_primitive_number() {
    assert_eq!(
        run_js(
            r#"
try {
    throw 42;
} catch (n) {
    console.log(n);
    console.log(typeof n);
}
"#
        ),
        vec!["42", "number"]
    );
}

#[test]
fn throw_primitive_bigint() {
    assert_eq!(
        run_js(
            r#"
try {
    throw 9007199254740991n;
} catch (b) {
    console.log(typeof b);
    console.log(b);
}
"#
        ),
        vec!["bigint", "9007199254740991"]
    );
}

