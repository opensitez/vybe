// vybe-test: js/json_semantics_matrix/json_parse_array_root_can_be_reduced_by_reviver
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

const result = JSON.parse('[1,2,3]', (key, value) => {
    return key === "" ? value.length : value;
});
__check(__line(result), "3");
