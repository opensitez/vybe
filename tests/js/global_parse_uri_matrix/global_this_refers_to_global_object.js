// vybe-test: js/global_parse_uri_matrix/global_this_refers_to_global_object
// origin: languages/js/tests/js/test_global_parse_uri_matrix.rs

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

globalThis.__temp = 42;
__check(__line(__temp), "42");
delete globalThis.__temp;
