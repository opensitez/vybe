// vybe-test: js/temporal_api/instant_compare
// origin: languages/js/tests/js/test_temporal_api.rs

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

function cmp(x, y) { return x < y ? -1 : x > y ? 1 : 0; }
const a = 1000;
const b = 2000;
__check(__line(cmp(a, b)), "-1");
