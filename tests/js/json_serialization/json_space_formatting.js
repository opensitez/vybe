// vybe-test: js/json_serialization/json_space_formatting
// origin: languages/js/tests/js/test_json_serialization.rs

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

const obj = { a: 1, b: [2, 3] };
const pretty = JSON.stringify(obj, null, 2);
const lines = pretty.split("\n");
__check(__line(lines.length > 1), "true");
__check(__line(lines[0]), "{");
