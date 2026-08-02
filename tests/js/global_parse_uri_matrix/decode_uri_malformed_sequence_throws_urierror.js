// vybe-test: js/global_parse_uri_matrix/decode_uri_malformed_sequence_throws_urierror
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

try {
  decodeURI("%E0%A4%A");
  console.log("no error");
} catch (error) {
  console.log(error instanceof URIError);
}
