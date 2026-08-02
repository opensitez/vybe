// vybe-test: js/bigint_bitwise_operations_masks/test_js_bigint_64bit_masking_and_extraction
// origin: languages/js/tests/js/test_js_bigint_bitwise_operations_masks.rs

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

const packed = (0x12345678n << 32n) | 0x9ABCDEF0n;
const high = (packed >> 32n) & 0xFFFFFFFFn;
const low = packed & 0xFFFFFFFFn;
__check(__line(high.toString(16) + "|" + low.toString(16)), "12345678|9abcdef0");
