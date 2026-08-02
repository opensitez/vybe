// vybe-test: js/symbol_advanced/symbol_has_instance_custom
// origin: languages/js/tests/js/test_symbol_advanced.rs

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

class OddNumbers {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 !== 0;
    }
}
__check(__line(1 instanceof OddNumbers), "true");
__check(__line(2 instanceof OddNumbers), "false");
__check(__line(3 instanceof OddNumbers), "true");
