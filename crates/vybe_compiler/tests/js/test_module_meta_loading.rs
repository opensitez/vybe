crate::js_cases! {
    import_meta_is_object => {
        r#"
console.log(typeof import.meta);
"#,
        ["object"]
    };

    import_meta_url_is_string => {
        r#"
console.log(typeof import.meta.url);
"#,
        ["string"]
    };

    import_meta_accepts_custom_properties => {
        r#"
import.meta.build = "canary";
console.log(import.meta.build);
"#,
        ["canary"]
    };

    top_level_await_resolves_promise_value => {
        r#"
const value = await Promise.resolve(41);
console.log(value + 1);
"#,
        ["42"]
    };

    top_level_await_preserves_statement_order => {
        r#"
console.log("before");
const value = await Promise.resolve("after await");
console.log(value);
"#,
        ["before", "after await"]
    };
}

crate::js_import_cases! {
    dynamic_import_returns_namespace_object => {
        r#"
const ns = await import("wasi:cli");
console.log(typeof ns);
console.log(typeof ns.log);
"#,
        ["object", "function"]
    };

    dynamic_import_exposes_namespace_keys => {
        r#"
const ns = await import("wasi:cli");
console.log(Object.keys(ns).includes("log"));
"#,
        ["true"]
    };

    dynamic_import_omits_default_for_host_namespace => {
        r#"
const ns = await import("wasi:cli");
console.log("default" in ns);
"#,
        ["false"]
    };
}
