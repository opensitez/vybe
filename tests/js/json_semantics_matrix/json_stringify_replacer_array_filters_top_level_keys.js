// vybe-test: js/json_semantics_matrix/json_stringify_replacer_array_filters_top_level_keys
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

__check(__line(JSON.stringify({ a: 1, b: 2, c: 3 }, ["c", "a"])), "{\"c\":3,\"a\":1}");
