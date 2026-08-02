// vybe-test: js/url_component_more_matrix/url_searchparams_set_encodes_spaces_as_plus_in_search
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

const u = new URL("https://example.com/");
u.searchParams.set("q", "two words");
__check(__line(u.search), "?q=two+words");
