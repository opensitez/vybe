// vybe-test: js/json_semantics_matrix/json_stringify_replacer_function_runs_after_tojson
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
    toJSON() {
        return { x: 2 };
    }
};
__check(__line(JSON.stringify(obj, (key, value) => {
    return typeof value === "number" ? value * 2 : value;
})), "{\"x\":4}");
