//! Error.prototype.toString, name/message properties, subclass toString, cause.

crate::js_cases! {
    builtin_error_tostring_with_message => {
        r#"console.log(new Error("boom").toString());"#,
        ["Error: boom"]
    };

    builtin_error_tostring_empty_message_omits_colon => {
        r#"console.log(new Error("").toString());"#,
        ["Error"]
    };

    builtin_typeerror_tostring_includes_name => {
        r#"console.log(new TypeError("bad").toString());"#,
        ["TypeError: bad"]
    };

    custom_error_subclass_tostring_uses_subclass_name => {
        r#"class AppError extends Error {} console.log(new AppError("app").toString());"#,
        ["AppError: app"]
    };

    error_name_assignment_changes_tostring_prefix => {
        r#"const e=new Error("msg");e.name="Custom";console.log(e.toString());"#,
        ["Custom: msg"]
    };

    error_message_assignment_updates_tostring => {
        r#"const e=new Error("a");e.message="b";console.log(e.toString());"#,
        ["Error: b"]
    };

    error_cause_property_readable => {
        r#"const root=new Error("root");const e=new Error("wrap",{cause:root});console.log(e.cause.message);"#,
        ["root"]
    };

    error_prototype_is_object_prototype => {
        r#"console.log(Object.getPrototypeOf(Error.prototype)===Object.prototype);"#,
        ["true"]
    };

    error_instanceof_error => {
        r#"console.log(new TypeError() instanceof Error);"#,
        ["true"]
    };

    subclass_instanceof_base_and_constructor => {
        r#"class E extends RangeError {} const e=new E();console.log(e instanceof RangeError);console.log(e instanceof Error);"#,
        ["true", "true"]
    };

    error_name_defaults_match_constructor => {
        r#"console.log(new ReferenceError().name);"#,
        ["ReferenceError"]
    };

    error_message_defaults_empty => {
        r#"console.log(new SyntaxError().message);"#,
        [""]
    };

    aggregate_error_tostring_includes_message => {
        r#"console.log(new AggregateError([new Error("a")],"many").toString().includes("many"));"#,
        ["true"]
    };

    aggregate_error_errors_array_length => {
        r#"console.log(new AggregateError([1,2,3],"x").errors.length);"#,
        ["3"]
    };

    error_stack_is_string_when_present => {
        r#"const e=new Error("s");console.log(typeof e.stack);"#,
        ["string"]
    };

    error_to_string_on_subclass_with_empty_name => {
        r#"class Silent extends Error {} const e=new Silent("m");e.name="";console.log(e.toString());"#,
        [": m"]
    };

    thrown_error_name_preserved_in_catch => {
        r#"try{throw new URIError("bad");}catch(e){console.log(e.name);}"#,
        ["URIError"]
    };

    error_constructor_without_new_works => {
        r#"console.log(Error("x").message);"#,
        ["x"]
    };

    typeerror_constructor_without_new => {
        r#"console.log(TypeError("t").name);"#,
        ["TypeError"]
    };

    error_prototype_tostring_is_function => {
        r#"console.log(typeof Error.prototype.toString);"#,
        ["function"]
    };
}
