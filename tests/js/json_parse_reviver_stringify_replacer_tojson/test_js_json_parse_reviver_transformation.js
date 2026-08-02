// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_parse_reviver_transformation
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

const json = '{"date":"2026-07-22T00:00:00.000Z","val":10}';
const parsed = JSON.parse(json, (key, value) => {
    if (key === "date") return new Date(value);
    if (typeof value === "number") return value * 2;
    return value;
});
__check(__line((parsed.date instanceof Date) + "|" + parsed.val), "true|20");
