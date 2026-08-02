// vybe-test: js/ecma_operators/bigints_mixed_with_numbers_throw_for_arithmetic
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

try {
    console.log(1n + 1);
} catch (e) {
    console.log(e.name);
}
console.log(String(8n / 3n));
console.log(String((-5n) % 3n));
