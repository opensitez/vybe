// vybe-test: js/global_builtin_edges/decode_uri_restores_encoded_url_text
// origin: languages/js/tests/js/test_global_builtin_edges.rs

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

__check(__line(decodeURI("https://example.com/a%20path?q=hello%20world#hash")), "https://example.com/a path?q=hello world#hash");
