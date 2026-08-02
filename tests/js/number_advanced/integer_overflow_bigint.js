// vybe-test: js/number_advanced/integer_overflow_bigint
// origin: languages/js/tests/js/test_number_advanced.rs

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

const MAX_SAFE = Number.MAX_SAFE_INTEGER;
__check(__line(MAX_SAFE + 1 === MAX_SAFE + 2), "true");  // loses precision
const bigMax = BigInt(MAX_SAFE);
__check(__line(bigMax + 1n === bigMax + 2n), "false");    // precise
__check(__line((bigMax + 1n).toString()), "9007199254740992");
