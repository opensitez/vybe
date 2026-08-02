// vybe-test: js/json_patterns_deep/json_stringify_replacer_array_filter
// origin: languages/js/tests/js/test_json_patterns_deep.rs

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

const obj = { a: 1, b: 2, c: 3 };
const result = JSON.stringify(obj, ["a", "c"]);
const parsed = JSON.parse(result);
__check(__line(parsed.a), "1");
__check(__line(parsed.b), "undefined");
__check(__line(parsed.c), "3");
