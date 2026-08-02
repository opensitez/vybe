// vybe-test: js/symbol_advanced/symbol_species_in_subclass
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

class MyArray extends Array {
    static get [Symbol.species]() { return Array; }
}
const m = new MyArray(1, 2, 3);
const mapped = m.map(x => x * 2);
__check(__line(mapped instanceof Array), "true");
__check(__line(mapped instanceof MyArray), "false"); // false due to Symbol.species
