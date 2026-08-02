// vybe-test: js/number_advanced/integer_division_floor_trunc_sign
// origin: languages/js/tests/js/test_number_advanced.rs

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

// Floor division (Python-style)
const floorDiv = (a, b) => Math.floor(a / b);
// Truncation division (C-style)
const truncDiv = (a, b) => Math.trunc(a / b);
__check(__line(floorDiv(7, 2)), "3");
__check(__line(floorDiv(-7, 2)), "-4");
__check(__line(truncDiv(-7, 2)), "-3");
__check(__line(floorDiv(7, -2)), "-4");
