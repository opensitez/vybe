// vybe-test: js/json_semantics_matrix/json_stringify_tojson_precedes_replacer_observation
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
    inner: {
        toJSON() {
            return { x: 1 };
        }
    }
};
JSON.stringify(obj, (key, value) => {
    if (key === "inner") {
        __check(__line(value.x), "1");
    }
    return value;
});
