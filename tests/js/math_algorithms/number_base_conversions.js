// vybe-test: js/math_algorithms/number_base_conversions
// origin: languages/js/tests/js/test_math_algorithms.rs

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

const dec = n => parseInt(n, 2);  // binary to decimal
const hex = n => n.toString(16);
const bin = n => n.toString(2);
__check(__line(dec("1010")), "10");
__check(__line(hex(255)), "ff");
__check(__line(bin(42)), "101010");
__check(__line(parseInt("ff", 16)), "255");
