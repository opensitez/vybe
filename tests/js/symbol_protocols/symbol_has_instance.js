// vybe-test: js/symbol_protocols/symbol_has_instance
// origin: languages/js/tests/js/test_symbol_protocols.rs

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

class EvenNumber {
    static [Symbol.hasInstance](n) {
        return typeof n === "number" && n % 2 === 0;
    }
}
__check(__line(2 instanceof EvenNumber), "true");
__check(__line(3 instanceof EvenNumber), "false");
__check(__line(100 instanceof EvenNumber), "true");
