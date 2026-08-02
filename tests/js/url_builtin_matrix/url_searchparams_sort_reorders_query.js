// vybe-test: js/url_builtin_matrix/url_searchparams_sort_reorders_query
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

const u = new URL("https://example.com/a?z=1&a=2&m=3");
u.searchParams.sort();
__check(__line(u.search), "?a=2&m=3&z=1");
