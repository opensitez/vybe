// vybe-test: js/number_bigint/number_to_fixed_various_decimals
// origin: languages/js/tests/js/test_number_bigint.rs

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

let n = 1.23456;
__check(__line(n.toFixed(0)), "1");
__check(__line(n.toFixed(2)), "1.23");
__check(__line(n.toFixed(4)), "1.2346");
