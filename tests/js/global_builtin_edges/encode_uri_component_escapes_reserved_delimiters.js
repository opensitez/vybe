// vybe-test: js/global_builtin_edges/encode_uri_component_escapes_reserved_delimiters
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

__check(__line(encodeURIComponent("a/b?c=d e")), "a%2Fb%3Fc%3Dd%20e");
