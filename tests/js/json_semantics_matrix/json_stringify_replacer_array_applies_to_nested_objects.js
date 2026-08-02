// vybe-test: js/json_semantics_matrix/json_stringify_replacer_array_applies_to_nested_objects
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

const obj = { outer: { a: 1, b: 2 }, a: 9 };
__check(__line(JSON.stringify(obj, ["outer", "b"])), "{\"outer\":{\"b\":2}}");
