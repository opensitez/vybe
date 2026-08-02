// vybe-test: js/class_patterns/symbol_hasinstance
// origin: languages/js/tests/js/test_class_patterns.rs

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

class Even {
    static [Symbol.hasInstance](num) {
        return typeof num === "number" && num % 2 === 0;
    }
}
__check(__line(4 instanceof Even), "true");
__check(__line(3 instanceof Even), "false");
