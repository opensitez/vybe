// vybe-test: js/url_builtin_matrix/url_setting_href_reparses_components
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

const u = new URL("https://example.com/a");
u.href = "http://user:pass@other.test:8080/b?q=1#h";
__check(__line(u.protocol), "http:");
__check(__line(u.host), "other.test:8080");
__check(__line(u.pathname), "/b");
