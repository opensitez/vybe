// vybe-test: js/bigint_advanced/bigint_throws_on_mixed_arithmetic_with_number
// origin: languages/js/tests/js/test_bigint_advanced.rs

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

let threw = false;
try { 1n + 1; } catch { threw = true; }
__check(__line(threw), "true");
