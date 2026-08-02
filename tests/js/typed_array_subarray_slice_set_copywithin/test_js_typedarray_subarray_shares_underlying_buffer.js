// vybe-test: js/typed_array_subarray_slice_set_copywithin/test_js_typedarray_subarray_shares_underlying_buffer
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
const sub = original.subarray(1, 3);
sub[0] = 99; // Modifying subarray updates original array!

__check(__line(original.join(",") + "|subLen=" + sub.length), "10,99,30,40|subLen=2");
