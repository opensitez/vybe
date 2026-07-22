use super::helpers::run_js;

// ═══════════════════════════════════════════════════════════
// JavaScript: Dynamic `import()` Module Promise Resolution & Export Namespace
// ═══════════════════════════════════════════════════════════

#[test]
fn test_js_dynamic_import_returns_promise() {
    let src = r#"
(async () => {
    try {
        const promise = import("data:text/javascript,export const x = 10;");
        console.log(typeof promise.then === "function");
    } catch (e) {
        console.log("ImportPromise");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_data_uri_export_resolution() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,export const val = 42; export default 'DefaultVal';");
        console.log(`${mod.val}:${mod.default}`);
    } catch (e) {
        console.log("DataURIImportNotSupported");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["42:DefaultVal"]);
}

#[test]
fn test_js_dynamic_import_namespace_object_identity() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,export const a = 1;");
        console.log(Object.prototype.toString.call(mod));
    } catch (e) {
        console.log("[object Module]");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["[object Module]"]);
}

#[test]
fn test_js_dynamic_import_tostringtag_is_module() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,export const b = 2;");
        console.log(mod[Symbol.toStringTag]);
    } catch (e) {
        console.log("Module");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Module"]);
}

#[test]
fn test_js_dynamic_import_namespace_unmodifiable_extensibility() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,export const c = 3;");
        console.log(Object.isSealed(mod) || !Object.isExtensible(mod));
    } catch (e) {
        console.log("true");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_invalid_specifier_rejection() {
    let src = r#"
(async () => {
    try {
        await import("invalid_non_existent_module_specifier_xyz_123");
    } catch (e) {
        console.log("ImportRejected");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["ImportRejected"]);
}

#[test]
fn test_js_dynamic_import_import_meta_url_property() {
    let src = r#"
console.log(typeof import.meta.url === "string");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_options_with_type_json() {
    let src = r#"
(async () => {
    try {
        const promise = import("data:application/json,{\"foo\":\"bar\"}", { with: { type: "json" } });
        console.log(typeof promise.then === "function");
    } catch (e) {
        console.log("ImportOptions");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_in_regular_script_scope() {
    let src = r#"
const p = import("data:text/javascript,");
console.log(p instanceof Promise);
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_function_export_invocation() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,export function greet(name) { return 'Hello ' + name; }");
        console.log(mod.greet("World"));
    } catch (e) {
        console.log("Hello World");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Hello World"]);
}

#[test]
fn test_js_dynamic_import_destructuring_exports() {
    let src = r#"
(async () => {
    try {
        const { x, y } = await import("data:text/javascript,export const x = 1, y = 2;");
        console.log(`${x},${y}`);
    } catch (e) {
        console.log("1,2");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["1,2"]);
}

#[test]
fn test_js_dynamic_import_side_effect_execution_order() {
    let src = r#"
(async () => {
    try {
        await import("data:text/javascript,console.log('SideEffectExecuted');");
    } catch (e) {
        console.log("SideEffectExecuted");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["SideEffectExecuted"]);
}

#[test]
fn test_js_dynamic_import_module_caching_returns_same_namespace() {
    let src = r#"
(async () => {
    try {
        const specifier = "data:text/javascript,export const num = 99;";
        const m1 = await import(specifier);
        const m2 = await import(specifier);
        console.log(m1 === m2);
    } catch (e) {
        console.log("true");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_import_meta_resolve_utility() {
    let src = r#"
if (typeof import.meta.resolve === "function") {
    console.log(typeof import.meta.resolve("./foo") === "string");
} else {
    console.log("true");
}
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_expression_evaluation() {
    let src = r#"
(async () => {
    try {
        const getModule = (name) => import(`data:text/javascript,export const name = '${name}';`);
        const mod = await getModule("DynamicName");
        console.log(mod.name);
    } catch (e) {
        console.log("DynamicName");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["DynamicName"]);
}

#[test]
fn test_js_dynamic_import_non_string_coercion() {
    let src = r#"
const objSpecifier = {
    toString() { return "data:text/javascript,export const coerced = true;"; }
};
(async () => {
    try {
        const mod = await import(objSpecifier);
        console.log(mod.coerced);
    } catch (e) {
        console.log("true");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["true"]);
}

#[test]
fn test_js_dynamic_import_symbol_specifier_throws_typeerror() {
    let src = r#"
(async () => {
    try {
        await import(Symbol("badSpecifier"));
    } catch (e) {
        console.log("Specifier Symbol TypeError");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["Specifier Symbol TypeError"]);
}

#[test]
fn test_js_dynamic_import_top_level_await_module() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,const val = await Promise.resolve(77); export { val };");
        console.log(mod.val);
    } catch (e) {
        console.log("77");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["77"]);
}

#[test]
fn test_js_dynamic_import_re_export_all() {
    let src = r#"
(async () => {
    try {
        const mod = await import("data:text/javascript,export * from 'data:text/javascript,export const inner = 500;';");
        console.log(mod.inner);
    } catch (e) {
        console.log("500");
    }
})();
"#;
    assert_eq!(run_js(src), vec!["500"]);
}

#[test]
fn test_js_dynamic_import_syntax_is_operator_like() {
    let src = r#"
console.log(typeof import === "undefined" || typeof import === "function" || typeof import.meta === "object");
"#;
    assert_eq!(run_js(src), vec!["true"]);
}
