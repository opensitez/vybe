// vybe-test: js/url_component_more_matrix/url_search_setter_preserves_hash
// origin: languages/js/tests/js/test_url_component_more_matrix.rs

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

const u = new URL("https://example.com/a#x");
u.search = "?q=1";
__check(__line(u.href), "https://example.com/a?q=1#x");
