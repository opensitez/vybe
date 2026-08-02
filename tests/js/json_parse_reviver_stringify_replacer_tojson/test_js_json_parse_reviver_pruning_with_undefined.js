// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_parse_reviver_pruning_with_undefined
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

const json = '{"keep":1,"remove":2}';
const parsed = JSON.parse(json, (key, value) => {
    if (key === "remove") return undefined; // returning undefined deletes property!
    return value;
});
__check(__line(parsed.keep + "|hasRemove=" + Object.hasOwn(parsed, "remove")), "1|hasRemove=false");
