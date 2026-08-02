// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_stringify_date_object_tojson_iso_format
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

const d = new Date(Date.UTC(2026, 6, 22));
__check(__line(JSON.stringify({ date: d })), "{\"date\":\"2026-07-22T00:00:00.000Z\"}");
