// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_parse_basic_object
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

const parsed = JSON.parse('{"a":1,"b":"hello","c":true}');
__check(__line(`${parsed.a}:${parsed.b}:${parsed.c}`), "1:hello:true");
