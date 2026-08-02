// vybe-test: js/urlsearchparams_edge_matrix/urlsearchparams_delete_missing_key_is_noop
// origin: languages/js/tests/js/test_urlsearchparams_edge_matrix.rs

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

const p = new URLSearchParams("a=1");
p.delete("x");
__check(__line(p.toString()), "a=1");
