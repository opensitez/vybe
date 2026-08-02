// vybe-test: js/global_parse_uri_matrix/encode_uri_component_encodes_slashes_and_question_mark
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

__check(__line(encodeURIComponent("a/b?c=d")), "a%2Fb%3Fc%3Dd");
