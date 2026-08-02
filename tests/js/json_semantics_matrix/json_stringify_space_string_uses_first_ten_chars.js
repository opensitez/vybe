// vybe-test: js/json_semantics_matrix/json_stringify_space_string_uses_first_ten_chars
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

const json = JSON.stringify({ a: { b: 1 } }, null, "abcdefghijklm");
__check(__line(json.indexOf('\nabcdefghijabcdefghij"b"') >= 0), "true");
