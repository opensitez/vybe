// vybe-test: js/symbol_has_instance_species_customization/test_js_symbol_species_custom_array_derived_type
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

class SpecialArray extends Array {
    static get [Symbol.species]() { return Array; }
}
const sa = new SpecialArray(1, 2, 3);
const mapped = sa.map(x => x * 2);
__check(__line(mapped.join(",") + "|isSpecial=" + (mapped instanceof SpecialArray) + "|isArray=" + (mapped instanceof Array)), "2,4,6|isSpecial=false|isArray=true");
