use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Nullish Coalescing (`??`) & Optional Chaining (`?.`) Combinations
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_optional_chaining_property_access_on_null_undefined() {
    let src = r#"
const obj = null;
console.log((obj?.prop === undefined) + "|" + (undefined?.nested?.val === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_optional_chaining_method_call_on_null_undefined() {
    let src = r#"
const user = { getAge: null };
console.log((user.getAge?.() === undefined) + "|" + (user.missingMethod?.() === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_optional_chaining_element_access_on_null_undefined() {
    let src = r#"
const arr = null;
console.log((arr?.[0] === undefined) + "|" + (undefined?.[10] === undefined));
"#;
    assert_eq!(run_js(src), vec!["true|true"]);
}

#[test]
fn test_js_optional_chaining_combined_with_nullish_coalescing() {
    let src = r#"
const user = { settings: null };
const theme = user.settings?.theme ?? "dark";
console.log(theme);
"#;
    assert_eq!(run_js(src), vec!["dark"]);
}

#[test]
fn test_js_nullish_coalescing_falsy_values_preservation() {
    let src = r#"
console.log(`${0 ?? 100}:${"" ?? "default"}:${false ?? true}:${NaN ?? 42}`);
"#;
    assert_eq!(run_js(src), vec!["0::false:NaN"]);
}

#[test]
fn test_js_nullish_coalescing_syntax_error_with_logical_operators() {
    let src = r#"
try {
    eval("true || false ?? 'default'"); // Mixing || or && with ?? without parentheses is a SyntaxError!
} catch (e) {
    console.log("Mixing Logical and Nullish SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Mixing Logical and Nullish SyntaxError"]);
}

#[test]
fn test_js_nullish_coalescing_parenthesized_with_logical_operators() {
    let src = r#"
console.log(((true || false) ?? "default") + "|" + ((null && true) ?? "fallback"));
"#;
    assert_eq!(run_js(src), vec!["true|fallback"]);
}

#[test]
fn test_js_optional_chaining_short_circuits_entire_chain() {
    let src = r#"
let sideEffectCount = 0;
const fn = () => { sideEffectCount++; return "data"; };
const obj = null;

const res = obj?.prop[fn()];
console.log((res === undefined) + "|SideEffects=" + sideEffectCount);
"#;
    assert_eq!(run_js(src), vec!["true|SideEffects=0"]);
}

#[test]
fn test_js_optional_chaining_deleting_property() {
    let src = r#"
const obj = { a: 1 };
const nullObj = null;
delete obj?.a;
delete nullObj?.b;
console.log(("a" in obj) + "|" + (nullObj === null));
"#;
    assert_eq!(run_js(src), vec!["false|true"]);
}

#[test]
fn test_js_optional_chaining_private_field_access() {
    let src = r#"
class Secret {
    #code = 1234;
    getCode(obj) {
        return obj?.#code;
    }
}
const s = new Secret();
console.log(s.getCode(s) + "|" + (s.getCode(null) === undefined));
"#;
    assert_eq!(run_js(src), vec!["1234|true"]);
}

#[test]
fn test_js_optional_chaining_function_call_with_arguments() {
    let src = r#"
const adder = (a, b) => a + b;
const missing = null;
console.log(adder?.(10, 20) + "|" + (missing?.(10, 20) === undefined));
"#;
    assert_eq!(run_js(src), vec!["30|true"]);
}

#[test]
fn test_js_nullish_coalescing_right_side_short_circuited() {
    let src = r#"
let evaluated = false;
const val = 10 ?? (evaluated = true, 20);
console.log(val + "|Evaluated=" + evaluated);
"#;
    assert_eq!(run_js(src), vec!["10|Evaluated=false"]);
}

#[test]
fn test_js_optional_chaining_deep_nested_structure() {
    let src = r#"
const data = { a: { b: { c: { d: "FoundDeep" } } } };
console.log(data?.a?.b?.c?.d + "|" + (data?.a?.x?.c?.d === undefined));
"#;
    assert_eq!(run_js(src), vec!["FoundDeep|true"]);
}

#[test]
fn test_js_optional_chaining_in_constructor_invocation_prohibited() {
    let src = r#"
try {
    eval("new Date?.();"); // new Target?.() is a SyntaxError!
} catch (e) {
    console.log("Optional Chaining Constructor SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Optional Chaining Constructor SyntaxError"]
    );
}

#[test]
fn test_js_optional_chaining_super_property_access_prohibited() {
    let src = r#"
try {
    eval("class B { m() { super?.m(); } }"); // super?. is a SyntaxError!
} catch (e) {
    console.log("Optional Chaining Super SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Optional Chaining Super SyntaxError"]);
}

#[test]
fn test_js_nullish_coalescing_chained_fallbacks() {
    let src = r#"
const a = null;
const b = undefined;
const c = "ThirdChoice";
console.log(a ?? b ?? c);
"#;
    assert_eq!(run_js(src), vec!["ThirdChoice"]);
}

#[test]
fn test_js_optional_chaining_tagged_template_literal_prohibited() {
    let src = r#"
try {
    eval("const tag = null; tag?.`template`;"); // Optional chaining tagged template is a SyntaxError!
} catch (e) {
    console.log("Optional Tagged Template SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Optional Tagged Template SyntaxError"]);
}

#[test]
fn test_js_optional_chaining_with_computed_symbol_property() {
    let src = r#"
const sym = Symbol("key");
const obj = { [sym]: "SymbolValue" };
const missing = null;
console.log(obj?.[sym] + "|" + (missing?.[sym] === undefined));
"#;
    assert_eq!(run_js(src), vec!["SymbolValue|true"]);
}

#[test]
fn test_js_optional_chaining_bigint_and_boolean_targets() {
    let src = r#"
const b = 100n;
const flag = true;
console.log(b?.toString() + "|" + flag?.valueOf());
"#;
    assert_eq!(run_js(src), vec!["100|true"]);
}

#[test]
fn test_js_optional_chaining_number_literal_dot_disambiguation() {
    let src = r#"
console.log((1)?.toString() + "|" + (0?.toString() === undefined));
"#;
    assert_eq!(run_js(src), vec!["1|false"]);
}
