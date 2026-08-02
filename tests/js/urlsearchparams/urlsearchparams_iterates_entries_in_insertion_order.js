// vybe-test: js/urlsearchparams/urlsearchparams_iterates_entries_in_insertion_order
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

const params = new URLSearchParams("a=1&b=2&a=3");
__check(__line([...params.entries()].map(([k, v]) => k + ":" + v).join(",")), "a:1,b:2,a:3");
