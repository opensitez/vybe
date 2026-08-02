// vybe-test: js/number_bigint/number_to_string_with_radix_hex_and_binary
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

__check(__line((255).toString(16)), "ff");
__check(__line((10).toString(2)), "1010");
__check(__line((8).toString(8)), "10");
