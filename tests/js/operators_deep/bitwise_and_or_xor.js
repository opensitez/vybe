// vybe-test: js/operators_deep/bitwise_and_or_xor
// origin: languages/js/tests/js/test_operators_deep.rs

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

__check(__line(0b1010 & 0b1100), "8");  // AND: 0b1000 = 8
__check(__line(0b1010 | 0b1100), "14");  // OR:  0b1110 = 14
__check(__line(0b1010 ^ 0b1100), "6");  // XOR: 0b0110 = 6
