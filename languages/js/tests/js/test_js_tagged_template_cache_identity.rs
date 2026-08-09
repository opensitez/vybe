use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Tagged Template Literal Template Object Caching Identity
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_tagged_template_cache_identity_same_call_site() {
    let src = r#"
function tag(strings) {
    return strings;
}
function getTemplate() {
    return tag`Hello ${1}`;
}
const t1 = getTemplate();
const t2 = getTemplate();
console.log(t1 === t2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_cache_identity_different_call_sites() {
    let src = r#"
function tag(strings) { return strings; }
const t1 = tag`Same Text`;
const t2 = tag`Same Text`;
console.log(t1 === t2); // Per ES2018 spec, different call sites produce distinct template objects
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_tagged_template_raw_array_identity_matches_template() {
    let src = r#"
function tag(strings) { return strings; }
function getRaw() { return tag`Text ${1}`.raw; }
const r1 = getRaw();
const r2 = getRaw();
console.log(r1 === r2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_cache_per_eval_in_loop() {
    let src = r#"
function tag(strings) { return strings; }
const templates = [];
for (let i = 0; i < 3; i++) {
    templates.push(tag`LoopTemplate`);
}
console.log(templates[0] === templates[1] && templates[1] === templates[2]);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_object_is_frozen() {
    let src = r#"
function tag(strings) { return strings; }
const t = tag`Sample ${"val"}`;
console.log(Object.isFrozen(t) + "|" + Object.isFrozen(t.raw));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_tagged_template_mutation_attempt_throws_in_strict() {
    let src = r#"
function tag(strings) {
    "use strict";
    try {
        strings[0] = "Mutated";
    } catch (e) {
        console.log("Mutation TypeError");
    }
}
tag`Original`;
"#;
    assert_eq!(run_js(src), vec!["Mutation TypeError"]);
}

#[test]
fn test_js_tagged_template_cache_in_recursive_function() {
    let src = r#"
function tag(strings) { return strings; }
function recurse(n) {
    const template = tag`RecurseNode`;
    if (n <= 0) return [template];
    return [template, ...recurse(n - 1)];
}
const list = recurse(2);
console.log(list[0] === list[1] && list[1] === list[2]);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_cache_with_dynamic_interpolations() {
    let src = r#"
function tag(strings, ...values) { return strings; }
function getT(v) { return tag`Value: ${v}`; }
const t1 = getT("A");
const t2 = getT("B");
console.log(t1 === t2); // Same callsite -> identical template array reference!
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_has_own_property_raw() {
    let src = r#"
function tag(strings) { return Object.hasOwn(strings, "raw"); }
console.log(tag`CheckRaw`);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_raw_property_non_enumerable() {
    let src = r#"
function tag(strings) {
    const desc = Object.getOwnPropertyDescriptor(strings, "raw");
    return desc.enumerable + "|" + desc.writable + "|" + desc.configurable;
}
console.log(tag`DescTest`);
"#;
    assert_eq!(run_js(src), vec!["false|false|false"]);
}

#[test]
fn test_js_tagged_template_cache_in_arrow_function_body() {
    let src = r#"
const tag = strings => strings;
const getFn = () => tag`ArrowTemplate`;
console.log(getFn() === getFn());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_cache_in_generator_function() {
    let src = r#"
function tag(strings) { return strings; }
function* gen() {
    yield tag`GenTemplate`;
    yield tag`GenTemplate`;
}
const g = gen();
const t1 = g.next().value;
const t2 = g.next().value;
console.log(t1 === t2); // Different call sites inside generator yield distinct objects
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_tagged_template_cache_in_class_method() {
    let src = r#"
class Parser {
    tag(strings) { return strings; }
    parse() { return this.tag`ClassTemplate`; }
}
const p = new Parser();
console.log(p.parse() === p.parse());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_strings_length_matches_interpolations() {
    let src = r#"
function tag(strings, ...values) {
    return strings.length === values.length + 1;
}
console.log(tag`A ${1} B ${2} C`);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_strings_is_instanceof_array() {
    let src = r#"
function tag(strings) {
    return Array.isArray(strings) + "|" + (strings instanceof Array);
}
console.log(tag`CheckArray`);
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_tagged_template_cache_isolation_across_modules_simulation() {
    let src = r#"
function tag(strings) { return strings; }
const moduleA = () => tag`SharedText`;
const moduleB = () => tag`SharedText`;
console.log(moduleA() === moduleB()); // Different function declarations -> different call site identity
"#;
    assert_eq!(run_js(src), vec!["false"]);
}

#[test]
fn test_js_tagged_template_cache_nested_tagged_templates() {
    let src = r#"
function tag(strings) { return strings; }
function getNested() {
    return tag`Outer ${tag`Inner`}`;
}
const n1 = getNested();
const n2 = getNested();
console.log(n1 === n2);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_cache_identity_closure_rebinding() {
    let src = r#"
function makeTagger() {
    const tag = strings => strings;
    return () => tag`ClosureTemplate`;
}
const f1 = makeTagger();
const f2 = makeTagger();
console.log(f1() === f1());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_cache_identity_eval_code() {
    let src = r#"
function tag(strings) { return strings; }
const t1 = eval("tag`EvalTemplate`");
const t2 = eval("tag`EvalTemplate`");
console.log(t1 === t2);
"#;
    assert_eq!(run_js(src), vec!["false"]); // Each eval execution creates a distinct callsite
}

#[test]
fn test_js_tagged_template_cache_identity_constructor_method() {
    let src = r#"
function tag(strings) { return strings; }
class Item {
    constructor() {
        this.tpl = tag`CtorTemplate`;
    }
}
const i1 = new Item();
const i2 = new Item();
console.log(i1.tpl === i2.tpl);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_tagged_template_raw_array_length_matches_strings_length() {
    let src = r#"
function tag(strings) {
    return strings.length === strings.raw.length;
}
console.log(tag`A${1}B${2}C`);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
