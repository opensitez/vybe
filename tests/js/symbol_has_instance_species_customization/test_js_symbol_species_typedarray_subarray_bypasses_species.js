// vybe-test: js/symbol_has_instance_species_customization/test_js_symbol_species_typedarray_subarray_bypasses_species
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

class CustomUint8 extends Uint8Array {
    static get [Symbol.species]() { return Uint8Array; }
}
const cu8 = new CustomUint8([1, 2, 3]);
const sub = cu8.subarray(1); // TypedArray.prototype.subarray does NOT use Symbol.species!
__check(__line(sub instanceof CustomUint8), "true");
