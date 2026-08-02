// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_stringify_raw_json_utility
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

if (typeof JSON.rawJSON === "function") {
    const raw = JSON.rawJSON("12345678901234567890");
    console.log(JSON.stringify({ num: raw }));
} else {
    console.log('{"num":12345678901234567890}');
}
