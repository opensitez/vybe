// vybe-test: js/url_builtin_matrix/url_searchparams_delete_removes_key
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

const u = new URL("https://example.com/a?x=1&y=2&x=3");
u.searchParams.delete("x");
__check(__line(u.href), "https://example.com/a?y=2");
