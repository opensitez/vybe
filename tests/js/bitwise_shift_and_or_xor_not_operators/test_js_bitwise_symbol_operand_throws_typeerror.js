// vybe-test: js/bitwise_shift_and_or_xor_not_operators/test_js_bitwise_symbol_operand_throws_typeerror
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
    const res = Symbol("a") | 1;
} catch (e) {
    __check(__line("Bitwise Symbol TypeError"), "Bitwise Symbol TypeError");
}
