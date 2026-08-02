// vybe-test: js/typed_array_uint8_int32_float64_views/test_js_typedarray_symbol_species_override
// origin: languages/js/tests/js/test_js_typed_array_uint8_int32_float64_views.rs

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

class CustomUint8 extends Uint8Array {}
const cu8 = new CustomUint8([5, 10]);
const sliced = cu8.slice(0, 1);
__check(__line(sliced[0] + "|isCustom=" + (sliced instanceof CustomUint8)), "5|isCustom=true");
