// vybe-test: js/json_deep/stringify_nan_infinity_become_null
// origin: languages/js/tests/js/test_json_deep.rs

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

const obj = { a: NaN, b: Infinity, c: -Infinity };
const result = JSON.parse(JSON.stringify(obj));
__check(__line(result.a), "null");
__check(__line(result.b), "null");
__check(__line(result.c), "null");
