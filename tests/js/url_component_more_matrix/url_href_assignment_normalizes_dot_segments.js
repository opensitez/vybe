// vybe-test: js/url_component_more_matrix/url_href_assignment_normalizes_dot_segments
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

const u = new URL("https://example.com/a");
u.href = "https://example.com/a/../b";
__check(__line(u.href), "https://example.com/b");
