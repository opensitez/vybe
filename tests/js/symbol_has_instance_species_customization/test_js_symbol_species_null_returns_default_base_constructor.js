// vybe-test: js/symbol_has_instance_species_customization/test_js_symbol_species_null_returns_default_base_constructor
// origin: languages/js/tests/js/test_js_symbol_has_instance_species_customization.rs

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

class NullSpeciesArray extends Array {
    static get [Symbol.species]() { return null; }
}
const nsa = new NullSpeciesArray(10, 20);
const res = nsa.map(x => x);
__check(__line(res instanceof Array + "|" + (res instanceof NullSpeciesArray)), "true|false");
