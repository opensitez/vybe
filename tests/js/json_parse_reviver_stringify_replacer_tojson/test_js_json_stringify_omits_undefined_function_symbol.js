// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_stringify_omits_undefined_function_symbol
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

const obj = {
    u: undefined,
    f: () => {},
    s: Symbol("id"),
    valid: "ok"
};
__check(__line(JSON.stringify(obj)), "{\"valid\":\"ok\"}");
