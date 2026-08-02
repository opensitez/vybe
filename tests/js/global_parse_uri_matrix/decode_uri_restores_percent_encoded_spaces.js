// vybe-test: js/global_parse_uri_matrix/decode_uri_restores_percent_encoded_spaces
// origin: languages/js/tests/js/test_global_parse_uri_matrix.rs

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

__check(__line(decodeURI("https://example.com/a%20b/c%20d")), "https://example.com/a b/c d");
