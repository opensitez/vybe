// vybe-test: js/number_advanced/hex_octal_binary_ops
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

const hex = 0xFF;
const octal = 0o17;
const binary = 0b1010;
__check(__line(hex), "255");
__check(__line(octal), "15");
__check(__line(binary), "10");
__check(__line((hex & binary).toString(2)), "1010");
__check(__line((octal | binary).toString(2)), "1111");
