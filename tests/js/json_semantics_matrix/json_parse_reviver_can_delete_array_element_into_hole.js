// vybe-test: js/json_semantics_matrix/json_parse_reviver_can_delete_array_element_into_hole
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

const arr = JSON.parse('[1,2,3]', (key, value) => {
    return key === "1" ? undefined : value;
});
__check(__line(1 in arr), "false");
__check(__line(arr[1]), "undefined");
__check(__line(JSON.stringify(arr)), "[1,null,3]");
