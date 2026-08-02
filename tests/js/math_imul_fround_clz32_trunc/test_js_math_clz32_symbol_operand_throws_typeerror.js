// vybe-test: js/math_imul_fround_clz32_trunc/test_js_math_clz32_symbol_operand_throws_typeerror
// origin: languages/js/tests/js/test_js_math_imul_fround_clz32_trunc.rs

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

const errors = [];
try {
    Math.clz32(Symbol("val"));
} catch (e) {
    errors.push("clz32");
}
__check(__line(errors.join("|")), "clz32");
