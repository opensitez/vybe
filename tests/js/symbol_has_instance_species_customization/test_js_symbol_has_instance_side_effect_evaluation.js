// vybe-test: js/symbol_has_instance_species_customization/test_js_symbol_has_instance_side_effect_evaluation
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

let evaluated = false;
const trap = {
    [Symbol.hasInstance]() {
        evaluated = true;
        return true;
    }
};
__check(__line((100 instanceof trap) + "|Evaluated=" + evaluated), "true|Evaluated=true");
