// vybe-test: js/symbol_has_instance_species_customization/test_js_symbol_has_instance_object_literal
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

const IntegerType = {
    [Symbol.hasInstance](val) {
        return Number.isInteger(val);
    }
};
__check(__line((42 instanceof IntegerType) + "|" + (3.14 instanceof IntegerType)), "true|false");
