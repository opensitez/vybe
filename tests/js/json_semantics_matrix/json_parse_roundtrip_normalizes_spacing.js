// vybe-test: js/json_semantics_matrix/json_parse_roundtrip_normalizes_spacing
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

const compact = JSON.stringify(JSON.parse('{ "a" : [1, 2], "b" : true }'));
__check(__line(compact), "{\"a\":[1,2],\"b\":true}");
