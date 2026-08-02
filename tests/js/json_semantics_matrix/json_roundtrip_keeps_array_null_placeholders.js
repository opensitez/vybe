// vybe-test: js/json_semantics_matrix/json_roundtrip_keeps_array_null_placeholders
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

const back = JSON.parse(JSON.stringify([1, null, 3]));
__check(__line(back[1] === null), "true");
