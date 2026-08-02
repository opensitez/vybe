// vybe-test: js/ecma_operators/bitwise_assign
// origin: languages/js/tests/js/test_ecma_operators.rs

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

let x = 0xFF;
x &= 0x0F;
__check(__line(x), "15");
x |= 0xF0;
__check(__line(x), "255");
x ^= 0xFF;
__check(__line(x), "0");
