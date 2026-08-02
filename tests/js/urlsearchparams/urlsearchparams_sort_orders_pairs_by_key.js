// vybe-test: js/urlsearchparams/urlsearchparams_sort_orders_pairs_by_key
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

const params = new URLSearchParams("z=1&a=2&m=3");
params.sort();
__check(__line(params.toString()), "a=2&m=3&z=1");
