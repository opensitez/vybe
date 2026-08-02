// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_of_subclass_constructor_inheritance
// origin: languages/js/tests/js/test_js_typed_array_from_of_constructors.rs

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
const cu8 = CustomUint8.of(10, 20);
__check(__line(cu8.join(",") + "|isCustom=" + (cu8 instanceof CustomUint8)), "10,20|isCustom=true");
