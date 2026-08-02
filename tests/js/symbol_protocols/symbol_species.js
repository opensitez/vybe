// vybe-test: js/symbol_protocols/symbol_species
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

class MyArray extends Array {
    static get [Symbol.species]() { return Array; }
    sum() { return this.reduce((a, b) => a + b, 0); }
}
const ma = new MyArray();
ma.push(1, 2, 3, 4);
const mapped = ma.map(x => x * 2);
// With Symbol.species = Array, map returns a plain Array
__check(__line(mapped instanceof Array), "true");
__check(__line(ma instanceof MyArray), "true");
__check(__line(ma.sum()), "10");
