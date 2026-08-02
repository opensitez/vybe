// vybe-test: js/json_semantics_matrix/json_stringify_tojson_on_root_object_controls_output
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

const obj = {
    a: 1,
    toJSON() {
        return { b: 2 };
    }
};
__check(__line(JSON.stringify(obj)), "{\"b\":2}");
