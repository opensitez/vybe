use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// ECMAScript: Modules — import/export syntax
// These test grammar parsing and compilation of module syntax.
// Actual resolution requires filesystem; we test syntax works.
// ═══════════════════════════════════════════════════════════

#[test]
fn export_function() {
    // Just verifies export syntax parses and compiles
    let out = run_js(r#"
export function add(a, b) {
    return a + b;
}
console.log(add(1, 2));
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn export_const() {
    let out = run_js(r#"
export const PI = 3.14159;
console.log(PI);
"#);
    assert_eq!(out, vec!["3.14159"]);
}

#[test]
fn export_class() {
    let out = run_js(r#"
export class MyClass {
    constructor(x) { this.x = x; }
    get() { return this.x; }
}
const m = new MyClass(42);
console.log(m.get());
"#);
    assert_eq!(out, vec!["42"]);
}

#[test]
fn export_default_function() {
    let out = run_js(r#"
export default function greet(name) {
    return "Hello " + name;
}
console.log(greet("World"));
"#);
    assert_eq!(out, vec!["Hello World"]);
}

#[test]
fn export_default_class() {
    let out = run_js(r#"
export default class {
    speak() { return "woof"; }
}
"#);
    // Just verify it parses — anonymous default class
    assert!(out.is_empty());
}

#[test]
fn export_named_list() {
    let out = run_js(r#"
const a = 1;
const b = 2;
export { a, b };
console.log(a + b);
"#);
    assert_eq!(out, vec!["3"]);
}

#[test]
fn export_with_alias() {
    let out = run_js(r#"
const x = 42;
export { x as answer };
console.log(x);
"#);
    assert_eq!(out, vec!["42"]);
}
