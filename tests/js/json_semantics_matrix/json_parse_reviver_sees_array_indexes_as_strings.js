// vybe-test: js/json_semantics_matrix/json_parse_reviver_sees_array_indexes_as_strings
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

JSON.parse('[10]', (key, value) => {
    if (key !== "") {
        __check(__line(typeof key + ":" + key), "string:0");
    }
    return value;
});
