// vybe-test: js/number_advanced/numeric_types_detection
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

const isInt = n => Number.isInteger(n);
const isFloat = n => typeof n === "number" && !Number.isInteger(n) && isFinite(n);
const isBigInt = n => typeof n === "bigint";
__check(__line(isInt(5)), "true");
__check(__line(isInt(5.0)), "true");
__check(__line(isFloat(5.5)), "true");
__check(__line(isFloat(Infinity)), "false");
__check(__line(isBigInt(42n)), "true");
