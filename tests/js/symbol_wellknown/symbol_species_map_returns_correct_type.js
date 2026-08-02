// vybe-test: js/symbol_wellknown/symbol_species_map_returns_correct_type
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

class PowerArray extends Array {
  static get [Symbol.species]() { return Array; }
}
const arr = new PowerArray(1, 2, 3);
const mapped = arr.map(x => x * 2);
__check(__line(mapped instanceof PowerArray), "false");
__check(__line(mapped instanceof Array), "true");
