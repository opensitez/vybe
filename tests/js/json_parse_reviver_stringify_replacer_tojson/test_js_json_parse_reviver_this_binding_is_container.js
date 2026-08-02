// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_parse_reviver_this_binding_is_container
// origin: languages/js/tests/js/test_js_json_parse_reviver_stringify_replacer_tojson.rs

function __line(...args) {
    // console.log joins its arguments with a single space. String() is the
    // coercion Vybe's logging host applies to each one.
    return args.map(String).join(" ");
}

function __check(got, want) {
    if (got !== want) {
        console.log("FAIL: want [" + want + "] got [" + got + "]");
        throw new Error("assertion failed");
    }
}

const json = '{"a":1}';
JSON.parse(json, function(key, value) {
    if (key === "a") {
        __check(__line(this.a), "1");
    }
    return value;
});
