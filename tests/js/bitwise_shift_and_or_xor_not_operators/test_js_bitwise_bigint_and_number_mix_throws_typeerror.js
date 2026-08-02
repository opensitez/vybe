// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_bigint_and_number_mix_throws_typeerror
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

try {
    const res = 10n & 5; // Cannot mix BigInt and Number in bitwise operations!
} catch (e) {
    __check(__line("Bitwise BigInt Number Mix TypeError"), "Bitwise BigInt Number Mix TypeError");
}
