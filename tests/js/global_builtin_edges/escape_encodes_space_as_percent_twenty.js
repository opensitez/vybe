// vybe-test: js/global_builtin_edges/escape_encodes_space_as_percent_twenty
// origin: languages/js/tests/js/test_global_builtin_edges.rs

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

__check(__line(escape("hello world")), "hello%20world");
