// vybe-test: js/urlsearchparams/urlsearchparams_constructs_from_sequence_pairs
// origin: languages/js/tests/js/test_urlsearchparams.rs

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

const params = new URLSearchParams([["x", "1"], ["y", "2"]]);
__check(__line(params.toString()), "x=1&y=2");
