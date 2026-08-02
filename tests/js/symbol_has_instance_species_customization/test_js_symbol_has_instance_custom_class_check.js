// vybe-test: js/symbol_has_instance_species_customization/test_js_symbol_has_instance_custom_class_check
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

class EvenNumber {
    static [Symbol.hasInstance](instance) {
        return typeof instance === "number" && instance % 2 === 0;
    }
}
__check(__line((2 instanceof EvenNumber) + "|" + (3 instanceof EvenNumber)), "true|false");
