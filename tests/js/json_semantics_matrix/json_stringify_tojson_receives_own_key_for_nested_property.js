// vybe-test: js/json_semantics_matrix/json_stringify_tojson_receives_own_key_for_nested_property
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

const outer = {
    inner: {
        toJSON(key) {
            __check(__line(key), "inner");
            return 1;
        }
    }
};
__check(__line(JSON.stringify(outer)), "{\"inner\":1}");
