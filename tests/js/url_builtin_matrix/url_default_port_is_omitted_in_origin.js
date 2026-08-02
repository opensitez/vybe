// vybe-test: js/url_builtin_matrix/url_default_port_is_omitted_in_origin
// origin: languages/js/tests/js/test_url_builtin_matrix.rs

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

const u = new URL("https://example.com:443/a");
__check(__line(u.origin), "https://example.com");
__check(__line(u.host), "example.com");
