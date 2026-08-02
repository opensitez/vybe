// vybe-test: js/spread_rest_advanced/object_rest_collect_remaining
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

const { a, b, ...rest } = { a: 1, b: 2, c: 3, d: 4 };
__check(__line(a, b), "1 2");
__check(__line(Object.keys(rest).sort().join(",")), "c,d");
