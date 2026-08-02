// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_slice_copies_underlying_buffer
// origin: languages/js/tests/js/test_js_typed_array_subarray_slice_set_copywithin.rs

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

const original = new Uint8Array([10, 20, 30, 40]);
const sliced = original.slice(1, 3);
sliced[0] = 99; // Modifying sliced copy does NOT update original!

__check(__line(original.join(",") + "|" + sliced.join(",")), "10,20,30,40|99,30");
