// vybe-test: js/json_semantics_matrix/json_stringify_replacer_function_can_scale_numbers
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

__check(__line(JSON.stringify({ a: 1, b: 2 }, (key, value) => {
    return typeof value === "number" ? value * 10 : value;
})), "{\"a\":10,\"b\":20}");
