// vybe-test: js/symbol_wellknown/symbol_hasinstance_custom_instanceof
// origin: languages/js/tests/js/test_symbol_wellknown.rs

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
  static [Symbol.hasInstance](val) {
    return typeof val === "number" && val % 2 === 0;
  }
}
__check(__line(4 instanceof EvenNumber), "true");
__check(__line(3 instanceof EvenNumber), "false");
