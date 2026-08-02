// vybe-test: js/bigint_advanced/bigint_basic_arithmetic
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

const a = 9007199254740993n; // beyond MAX_SAFE_INTEGER
const b = 1n;
__check(__line((a + b).toString()), "9007199254740994");
__check(__line((a - b).toString()), "9007199254740992");
__check(__line((a * 2n).toString()), "18014398509481986");
