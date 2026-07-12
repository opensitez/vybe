crate::js_cases! {
    eval_can_compute_expression_result => {
        r#"
console.log(eval("1 + 2 + 3"));
"#,
        ["6"]
    };

    eval_can_define_local_binding_in_current_scope => {
        r#"
eval("var fromEval = 42;");
console.log(fromEval);
"#,
        ["42"]
    };

    eval_can_return_object_literal_when_wrapped => {
        r#"
const obj = eval("({ answer: 42 })");
console.log(obj.answer);
"#,
        ["42"]
    };

    encode_uri_preserves_url_structure_characters => {
        r#"
console.log(encodeURI("https://example.com/a path?q=hello world#hash"));
"#,
        ["https://example.com/a%20path?q=hello%20world#hash"]
    };

    decode_uri_restores_encoded_url_text => {
        r#"
console.log(decodeURI("https://example.com/a%20path?q=hello%20world#hash"));
"#,
        ["https://example.com/a path?q=hello world#hash"]
    };

    encode_uri_component_escapes_reserved_delimiters => {
        r#"
console.log(encodeURIComponent("a/b?c=d e"));
"#,
        ["a%2Fb%3Fc%3Dd%20e"]
    };

    decode_uri_component_restores_reserved_delimiters => {
        r#"
console.log(decodeURIComponent("a%2Fb%3Fc%3Dd%20e"));
"#,
        ["a/b?c=d e"]
    };

    escape_encodes_space_as_percent_twenty => {
        r#"
console.log(escape("hello world"));
"#,
        ["hello%20world"]
    };

    unescape_decodes_percent_encoded_text => {
        r#"
console.log(unescape("hello%20world"));
"#,
        ["hello world"]
    };
}
