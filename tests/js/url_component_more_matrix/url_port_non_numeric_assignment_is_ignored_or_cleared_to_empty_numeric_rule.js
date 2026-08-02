// vybe-test: js/url_component_more_matrix/url_port_non_numeric_assignment_is_ignored_or_cleared_to_empty_numeric_rule
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

const u = new URL("https://example.com:8080/a");
u.port = "abc";
__check(__line(u.port === "8080" || u.port === ""), "true");
