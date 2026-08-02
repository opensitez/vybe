// vybe-test: js/json_semantics_matrix/json_stringify_replacer_function_sees_object_children_individually
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

const seen = [];
JSON.stringify({ a: 1, b: { c: 2 } }, (key, value) => {
    if (key !== "") {
        seen.push(key);
    }
    return value;
});
__check(__line(seen.join(",")), "a,b,c");
