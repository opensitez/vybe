// vybe-test: js/nullish_coalescing_and_optional_chaining_combinations/test_js_nullish_coalescing_syntax_error_with_logical_operators
// origin: languages/js/tests/js/test_js_nullish_coalescing_and_optional_chaining_combinations.rs

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
    eval("true || false ?? 'default'"); // Mixing || or && with ?? without parentheses is a SyntaxError!
} catch (e) {
    __check(__line("Mixing Logical and Nullish SyntaxError"), "Mixing Logical and Nullish SyntaxError");
}
