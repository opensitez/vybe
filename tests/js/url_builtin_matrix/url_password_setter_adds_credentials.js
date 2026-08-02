// vybe-test: js/url_builtin_matrix/url_password_setter_adds_credentials
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
u.username = "alice";
u.password = "secret";
__check(__line(u.href), "https://alice:secret@example.com/a");
