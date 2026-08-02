// vybe-test: js/urlsearchparams/urlsearchparams_append_preserves_multiple_values
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

const params = new URLSearchParams();
params.append("tag", "js");
params.append("tag", "tests");
__check(__line(params.getAll("tag").join(",")), "js,tests");
