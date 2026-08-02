// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_rgb_color_packing_unpacking
// origin: languages/js/tests/js/test_js_bitwise_shift_and_or_xor_not_operators.rs

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

const r = 255, g = 128, b = 64;
const color = (r << 16) | (g << 8) | b;
const outR = (color >> 16) & 0xFF;
const outG = (color >> 8) & 0xFF;
const outB = color & 0xFF;
__check(__line(`${outR}:${outG}:${outB}`), "255:128:64");
