// vybe-test: js/url_builtin_matrix/url_absolute_components_basic
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

const u = new URL("https://user:pass@example.com:8080/a/b?q=1#hash");
__check(__line(u.protocol), "https:");
__check(__line(u.username), "user");
__check(__line(u.password), "pass");
__check(__line(u.hostname), "example.com");
__check(__line(u.port), "8080");
__check(__line(u.pathname), "/a/b");
__check(__line(u.search), "?q=1");
__check(__line(u.hash), "#hash");
