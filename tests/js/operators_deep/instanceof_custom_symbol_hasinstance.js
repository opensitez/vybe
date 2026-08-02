// vybe-test: js/operators_deep/instanceof_custom_symbol_hasinstance
// origin: languages/js/tests/js/test_operators_deep.rs

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

class EvenChecker {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 === 0;
    }
}
__check(__line(2 instanceof EvenChecker), "true");
__check(__line(3 instanceof EvenChecker), "false");
__check(__line(4 instanceof EvenChecker), "true");
