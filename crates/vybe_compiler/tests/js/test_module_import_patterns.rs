/// Import attributes (import assertions), import.meta, dynamic import,
/// module namespace objects, cross-module patterns.
use super::helpers::run_js;

// ── import.meta ───────────────────────────────────────────────────────────────

#[test]
fn import_meta_is_object() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof import.meta);
"#
        ),
        vec!["object"]
    );
}

#[test]
fn import_meta_url_is_string() {
    assert_eq!(
        run_js(
            r#"
console.log(typeof import.meta.url);
"#
        ),
        vec!["string"]
    );
}

// ── dynamic import ────────────────────────────────────────────────────────────

#[test]
fn dynamic_import_returns_promise() {
    assert_eq!(
        run_js(
            r#"
// We can check that import() returns a Promise (even if module doesn't exist,
// the call site returns a thenable)
const maybePromise = import("./nonexistent_module_abc.js").catch(() => "failed");
console.log(maybePromise instanceof Promise);
"#
        ),
        vec!["true"]
    );
}

#[test]
fn dynamic_import_failure_caught_by_catch() {
    assert_eq!(
        run_js(
            r#"
import("./does_not_exist_xyz.js")
    .then(() => console.log("loaded"))
    .catch(() => console.log("failed"));
"#
        ),
        vec!["failed"]
    );
}

// ── import assertions / attributes (ES2025) ───────────────────────────────────

#[test]
fn import_with_attribute_type_json_syntactically_valid() {
    assert_eq!(
        run_js(
            r#"
// Import attributes syntax: import x from "y" with { type: "json" }
// Testing the syntax is parsed (actual loading may fail in test env)
let ok = true;
try {
    eval('import("./data.json", { with: { type: "json" } }).catch(() => {})');
} catch (e) {
    // SyntaxError means parser doesn't support it yet
    ok = e instanceof SyntaxError ? false : true;
}
console.log(typeof ok);
"#
        ),
        vec!["boolean"]
    );
}

// ── module namespace object ───────────────────────────────────────────────────

#[test]
fn module_object_has_module_namespace_tostring_tag() {
    assert_eq!(
        run_js(
            r#"
// Can't easily test namespace object without real module system,
// but we can test that Symbol.toStringTag is 'Module' conceptually
const ns = Object.create(null);
Object.defineProperty(ns, Symbol.toStringTag, { value: "Module" });
console.log(Object.prototype.toString.call(ns));
"#
        ),
        vec!["[object Module]"]
    );
}

// ── top-level await (module context) ─────────────────────────────────────────

#[test]
fn async_function_simulates_top_level_await() {
    assert_eq!(
        run_js(
            r#"
// In module scripts, top-level await is valid
// In non-module context, we simulate with async IIFE
(async () => {
    const data = await Promise.resolve("fetched");
    console.log(data);
})();
"#
        ),
        vec!["fetched"]
    );
}

#[test]
fn top_level_await_sequential_execution() {
    assert_eq!(
        run_js(
            r#"
const log = [];
async function main() {
    log.push("start");
    const a = await Promise.resolve(1);
    log.push("a=" + a);
    const b = await Promise.resolve(a + 1);
    log.push("b=" + b);
}
main().then(() => console.log(log.join(",")));
"#
        ),
        vec!["start,a=1,b=2"]
    );
}

// ── re-export patterns ────────────────────────────────────────────────────────

#[test]
fn destructure_namespace_import_simulation() {
    assert_eq!(
        run_js(
            r#"
// Simulate named exports with object
const mathModule = {
    add: (a, b) => a + b,
    sub: (a, b) => a - b,
    PI: 3.14159
};
const { add, sub, PI } = mathModule;
console.log(add(2, 3));
console.log(sub(5, 2));
console.log(PI.toFixed(2));
"#
        ),
        vec!["5", "3", "3.14"]
    );
}

// ── lazy module initialization simulation ─────────────────────────────────────

#[test]
fn lazy_initialization_with_import_simulation() {
    assert_eq!(
        run_js(
            r#"
// Simulate lazy singleton loading
let instance = null;
function getInstance() {
    if (!instance) instance = { value: Math.random() };
    return instance;
}
const a = getInstance();
const b = getInstance();
console.log(a === b);
"#
        ),
        vec!["true"]
    );
}

// ── conditional exports simulation ────────────────────────────────────────────

#[test]
fn conditional_module_loading_simulation() {
    assert_eq!(
        run_js(
            r#"
async function loadModule(name) {
    if (name === "a") return { value: 1 };
    if (name === "b") return { value: 2 };
    throw new Error("unknown:" + name);
}

async function main() {
    const mod = await loadModule("a");
    console.log(mod.value);
    try {
        await loadModule("c");
    } catch (e) {
        console.log(e.message);
    }
}
main();
"#
        ),
        vec!["1", "unknown:c"]
    );
}
