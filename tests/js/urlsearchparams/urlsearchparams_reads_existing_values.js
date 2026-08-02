// vybe-test: js/urlsearchparams/urlsearchparams_reads_existing_values
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

const params = new URLSearchParams("q=vybe&page=2");
__check(__line(params.get("q")), "vybe");
__check(__line(params.get("page")), "2");
