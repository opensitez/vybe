use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Async — async/await, Promises
// ═══════════════════════════════════════════════════════════

#[test]
fn async_function_returns_value() {
    let out = run_js(r#"
async function getValue() {
    return 42;
}
const p = getValue();
console.log(p);
"#);
    assert!(!out.is_empty());
}

#[test]
fn async_arrow_function() {
    let out = run_js(r#"
const fn = async () => "hello";
const p = fn();
console.log(p);
"#);
    assert!(!out.is_empty());
}

#[test]
fn async_function_declaration() {
    let out = run_js(r#"
async function process(x) {
    return x * 2;
}
const r = process(21);
console.log(r);
"#);
    assert!(!out.is_empty());
}

#[test]
fn async_class_method() {
    let out = run_js(r#"
class Service {
    async fetch(url) {
        return "data from " + url;
    }
}
const s = new Service();
const r = s.fetch("api/users");
console.log(r);
"#);
    assert!(!out.is_empty());
}

#[test]
fn await_expression() {
    let out = run_js(r#"
async function main() {
    const x = await 42;
    console.log(x);
}
main();
"#);
    // await on a non-promise should resolve immediately
    assert_eq!(out, vec!["42"]);
}
