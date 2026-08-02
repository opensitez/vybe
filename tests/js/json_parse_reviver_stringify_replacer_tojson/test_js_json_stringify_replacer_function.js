// vybe-test: js/json_parse_reviver_stringify_replacer_tojson/test_js_json_stringify_replacer_function
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

const obj = { name: "Alice", age: 30 };
const json = JSON.stringify(obj, (key, value) => {
    if (typeof value === "number") return value + 5;
    return value;
});
__check(__line(json), "{\"name\":\"Alice\",\"age\":35}");
