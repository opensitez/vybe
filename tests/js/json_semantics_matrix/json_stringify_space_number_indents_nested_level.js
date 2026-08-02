// vybe-test: js/json_semantics_matrix/json_stringify_space_number_indents_nested_level
// origin: languages/js/tests/js/test_json_semantics_matrix.rs

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

const json = JSON.stringify({ a: 1, b: { c: 2 } }, null, 2);
__check(__line(json.indexOf('\n  "b"') >= 0), "true");
__check(__line(json.indexOf('\n    "c"') >= 0), "true");
