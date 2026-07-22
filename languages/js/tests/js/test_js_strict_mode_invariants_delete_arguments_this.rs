use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Strict Mode Invariants (`"use strict"`, `delete`, `arguments`, `this`)
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_strict_mode_assignment_to_undeclared_variable_throws_referenceerror() {
    let src = r#"
try {
    eval("'use strict'; undeclaredVar = 10;");
} catch (e) {
    console.log("Strict Undeclared Variable ReferenceError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Strict Undeclared Variable ReferenceError"]
    );
}

#[test]
fn test_js_strict_mode_assignment_to_read_only_property_throws_typeerror() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 10, writable: false });
try {
    "use strict";
    obj.fixed = 20;
} catch (e) {
    console.log("Strict Read-Only Property TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Read-Only Property TypeError"]);
}

#[test]
fn test_js_strict_mode_assignment_to_getter_only_property_throws_typeerror() {
    let src = r#"
const obj = {
    get val() { return 5; }
};
try {
    "use strict";
    obj.val = 10;
} catch (e) {
    console.log("Strict Getter-Only Property TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Getter-Only Property TypeError"]);
}

#[test]
fn test_js_strict_mode_delete_unqualified_identifier_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; var x = 1; delete x;");
} catch (e) {
    console.log("Strict Delete Unqualified Identifier SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Strict Delete Unqualified Identifier SyntaxError"]
    );
}

#[test]
fn test_js_strict_mode_delete_non_configurable_property_throws_typeerror() {
    let src = r#"
const obj = {};
Object.defineProperty(obj, "fixed", { value: 1, configurable: false });
try {
    "use strict";
    delete obj.fixed;
} catch (e) {
    console.log("Strict Delete Non-Configurable Property TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Strict Delete Non-Configurable Property TypeError"]
    );
}

#[test]
fn test_js_strict_mode_duplicate_parameter_names_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; function fn(a, a) {}");
} catch (e) {
    console.log("Strict Duplicate Parameters SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Duplicate Parameters SyntaxError"]);
}

#[test]
fn test_js_strict_mode_arguments_unmapped_to_parameters() {
    let src = r#"
function fn(a) {
    "use strict";
    a = 100;
    return arguments[0]; // arguments[0] remains original passed value (not mapped to parameter 'a')!
}
console.log(fn(5));
"#;
    assert_eq!(run_js(src), vec!["5"]);
}

#[test]
fn test_js_strict_mode_arguments_callee_access_throws_typeerror() {
    let src = r#"
function fn() {
    "use strict";
    return arguments.callee;
}
try {
    fn();
} catch (e) {
    console.log("Strict arguments.callee TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict arguments.callee TypeError"]);
}

#[test]
fn test_js_strict_mode_arguments_caller_access_throws_typeerror() {
    let src = r#"
function fn() {
    "use strict";
    return arguments.caller;
}
try {
    fn();
} catch (e) {
    console.log("Strict arguments.caller TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict arguments.caller TypeError"]);
}

#[test]
fn test_js_strict_mode_this_undefined_in_standalone_function() {
    let src = r#"
function getThis() {
    "use strict";
    return this;
}
console.log(getThis() === undefined);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_strict_mode_this_primitive_values_not_boxed() {
    let src = r#"
function testThis() {
    "use strict";
    return typeof this;
}
console.log(testThis.call(123) + "|" + testThis.call("hello") + "|" + testThis.call(true));
"#;
    assert_eq!(run_js(src), vec!["number|string|boolean"]);
}

#[test]
fn test_js_strict_mode_eval_arguments_binding_restricted() {
    let src = r#"
try {
    eval("'use strict'; var eval = 10;");
} catch (e) {
    console.log("Strict eval Binding SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict eval Binding SyntaxError"]);
}

#[test]
fn test_js_strict_mode_reserved_keywords_as_identifiers_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; var let = 5;");
} catch (e) {
    console.log("Strict Reserved Word SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Reserved Word SyntaxError"]);
}

#[test]
fn test_js_strict_mode_octal_literal_prohibited_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; var num = 0123;");
} catch (e) {
    console.log("Strict Legacy Octal Literal SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Legacy Octal Literal SyntaxError"]);
}

#[test]
fn test_js_strict_mode_octal_escape_sequence_in_string_throws_syntaxerror() {
    let src = r#"
try {
    eval("'use strict'; var str = '\\123';");
} catch (e) {
    console.log("Strict Legacy Octal Escape SyntaxError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Legacy Octal Escape SyntaxError"]);
}

#[test]
fn test_js_strict_mode_class_body_is_implicitly_strict() {
    let src = r#"
class StrictClass {
    method() {
        try {
            undeclaredField = 10;
        } catch (e) {
            console.log("Implicit Class Strict ReferenceError");
        }
    }
}
new StrictClass().method();
"#;
    assert_eq!(run_js(src), vec!["Implicit Class Strict ReferenceError"]);
}

#[test]
fn test_js_strict_mode_es6_modules_are_implicitly_strict() {
    let src = r#"
function fnInModule() {
    try {
        eval("undeclaredInModule = 10;");
    } catch (e) {
        return true;
    }
    return false;
}
console.log(fnInModule());
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_strict_mode_function_caller_property_access_throws_typeerror() {
    let src = r#"
function fn() {
    "use strict";
    return fn.caller;
}
try {
    fn();
} catch (e) {
    console.log("Strict Function.caller TypeError");
}
"#;
    assert_eq!(run_js(src), vec!["Strict Function.caller TypeError"]);
}

#[test]
fn test_js_strict_mode_assignment_to_non_extensible_object_throws_typeerror() {
    let src = r#"
const obj = Object.preventExtensions({});
try {
    "use strict";
    obj.newProp = 10;
} catch (e) {
    console.log("Strict Non-Extensible Assignment TypeError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Strict Non-Extensible Assignment TypeError"]
    );
}

#[test]
fn test_js_strict_mode_directive_in_function_with_non_simple_parameters_throws() {
    let src = r#"
try {
    eval("function fn(a = 1) { 'use strict'; }"); // Non-simple parameters (defaults/destructuring) cannot have 'use strict'!
} catch (e) {
    console.log("Strict Directive Non-Simple Parameter SyntaxError");
}
"#;
    assert_eq!(
        run_js(src),
        vec!["Strict Directive Non-Simple Parameter SyntaxError"]
    );
}
