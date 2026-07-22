use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: `try...catch...finally` Return Override & Control Flow Unwinding
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_finally_return_overrides_try_return() {
    let src = r#"
function fn() {
    try {
        return "TryValue";
    } finally {
        return "FinallyValue";
    }
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["FinallyValue"]);
}

#[test]
fn test_js_finally_return_overrides_catch_return() {
    let src = r#"
function fn() {
    try {
        throw new Error("Err");
    } catch (e) {
        return "CatchValue";
    } finally {
        return "FinallyValue";
    }
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["FinallyValue"]);
}

#[test]
fn test_js_finally_throw_overrides_try_return() {
    let src = r#"
function fn() {
    try {
        return "TryValue";
    } finally {
        throw new Error("FinallyError");
    }
}
try {
    fn();
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["FinallyError"]);
}

#[test]
fn test_js_finally_throw_overrides_catch_throw() {
    let src = r#"
function fn() {
    try {
        throw new Error("TryError");
    } catch (e) {
        throw new Error("CatchError");
    } finally {
        throw new Error("FinallyError");
    }
}
try {
    fn();
} catch (e) {
    console.log(e.message);
}
"#;
    assert_eq!(run_js(src), vec!["FinallyError"]);
}

#[test]
fn test_js_finally_executes_even_if_no_catch_present() {
    let src = r#"
let finallyRun = false;
function fn() {
    try {
        return 10;
    } finally {
        finallyRun = true;
    }
}
console.log(fn() + "|FinallyRun=" + finallyRun);
"#;
    assert_eq!(run_js(src), vec!["10|FinallyRun=true"]);
}

#[test]
fn test_js_finally_executes_on_uncaught_exception_propagation() {
    let src = r#"
let finallyExecuted = false;
function fn() {
    try {
        throw new Error("Uncaught");
    } finally {
        finallyExecuted = true;
    }
}
try {
    fn();
} catch (e) {
    console.log(e.message + "|FinallyExecuted=" + finallyExecuted);
}
"#;
    assert_eq!(run_js(src), vec!["Uncaught|FinallyExecuted=true"]);
}

#[test]
fn test_js_catch_binding_destructuring() {
    let src = r#"
try {
    throw { code: 404, msg: "Not Found" };
} catch ({ code, msg }) {
    console.log(`${code}:${msg}`);
}
"#;
    assert_eq!(run_js(src), vec!["404:Not Found"]);
}

#[test]
fn test_js_optional_catch_binding_es2019() {
    let src = r#"
let caught = false;
try {
    throw new Error();
} catch {
    caught = true;
}
console.log(caught);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_try_catch_finally_nested_unwinding() {
    let src = r#"
const log = [];
function fn() {
    try {
        log.push("Try1");
        try {
            log.push("Try2");
            throw new Error("Err2");
        } catch (e) {
            log.push("Catch2");
            return "Ret2";
        } finally {
            log.push("Finally2");
        }
    } finally {
        log.push("Finally1");
    }
}
console.log(fn() + "|" + log.join(","));
"#;
    assert_eq!(run_js(src), vec!["Ret2|Try1,Try2,Catch2,Finally2,Finally1"]);
}

#[test]
fn test_js_catch_binding_scope_isolation() {
    let src = r#"
const err = "OuterErr";
try {
    throw "InnerErr";
} catch (err) {
    console.log("InsideCatch: " + err);
}
console.log("OutsideCatch: " + err);
"#;
    assert_eq!(
        run_js(src),
        vec!["InsideCatch: InnerErr", "OutsideCatch: OuterErr"]
    );
}

#[test]
fn test_js_catch_reassigning_error_param() {
    let src = r#"
try {
    throw "InitialErr";
} catch (e) {
    e = "ReassignedErr";
    console.log(e);
}
"#;
    assert_eq!(run_js(src), vec!["ReassignedErr"]);
}

#[test]
fn test_js_finally_suppresses_uncaught_error_with_return() {
    let src = r#"
function fn() {
    try {
        throw new Error("SuppressedError");
    } finally {
        return "SuppressedByReturn"; // Return in finally suppresses exception!
    }
}
console.log(fn());
"#;
    assert_eq!(run_js(src), vec!["SuppressedByReturn"]);
}

#[test]
fn test_js_async_try_catch_finally_unwinding() {
    let src = r#"
async function fn() {
    const log = [];
    try {
        log.push("AsyncTry");
        await Promise.reject(new Error("AsyncErr"));
    } catch (e) {
        log.push("AsyncCatch");
    } finally {
        log.push("AsyncFinally");
    }
    return log.join(",");
}
(async () => {
    console.log(await fn());
})();
"#;
    assert_eq!(run_js(src), vec!["AsyncTry,AsyncCatch,AsyncFinally"]);
}

#[test]
fn test_js_try_catch_generator_yield_in_finally() {
    let src = r#"
function* gen() {
    try {
        yield 1;
    } finally {
        yield 2;
    }
}
const g = gen();
console.log(`${g.next().value}:${g.next().value}:${g.next().done}`);
"#;
    assert_eq!(run_js(src), vec!["1:2:true"]);
}

#[test]
fn test_js_finally_return_expression_evaluated_before_cleanup() {
    let src = r#"
let sideEffect = 0;
function getVal() {
    sideEffect++;
    return sideEffect;
}
function fn() {
    try {
        return 99;
    } finally {
        return getVal();
    }
}
console.log(fn() + "|SideEffect=" + sideEffect);
"#;
    assert_eq!(run_js(src), vec!["1|SideEffect=1"]);
}

#[test]
fn test_js_try_statement_without_catch_or_finally_throws_syntaxerror() {
    let src = r#"
try {
    eval("try { const x = 1; }");
} catch (e) {
    console.log("Try Alone SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Try Alone SyntaxError"]);
}

#[test]
fn test_js_try_catch_finally_completion_value() {
    let src = r#"
console.log(eval("try { 10; } catch(e) {} finally { 20; }"));
"#;
    assert_eq!(run_js(src), vec!["20"]);
}

#[test]
fn test_js_catch_parameter_default_value_destructuring() {
    let src = r#"
try {
    throw {};
} catch ({ msg = "DefaultMsg" }) {
    console.log(msg);
}
"#;
    assert_eq!(run_js(src), vec!["DefaultMsg"]);
}

#[test]
fn test_js_try_finally_continue_loop() {
    let src = r#"
const log = [];
for (let i = 0; i < 2; i++) {
    try {
        log.push(`Try${i}`);
        continue;
    } finally {
        log.push(`Finally${i}`);
    }
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["Try0,Finally0,Try1,Finally1"]);
}

#[test]
fn test_js_try_finally_break_loop() {
    let src = r#"
const log = [];
for (let i = 0; i < 2; i++) {
    try {
        log.push(`Try${i}`);
        break;
    } finally {
        log.push(`Finally${i}`);
    }
}
console.log(log.join(","));
"#;
    assert_eq!(run_js(src), vec!["Try0,Finally0"]);
}
