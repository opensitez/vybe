// vybe-test: js/json_semantics_matrix/json_parse_reviver_can_wrap_nested_object_values
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

const obj = JSON.parse('{"box":{"value":2}}', (key, value) => {
    return key === "box" ? { wrapped: value.value + 1 } : value;
});
__check(__line(obj.box.wrapped), "3");
