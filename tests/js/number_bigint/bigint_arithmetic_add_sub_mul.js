// vybe-test: js/number_bigint/bigint_arithmetic_add_sub_mul
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

let a = 10n;
let b = 3n;
__check(__line(a + b), "13n");
__check(__line(a - b), "7n");
__check(__line(a * b), "30n");
