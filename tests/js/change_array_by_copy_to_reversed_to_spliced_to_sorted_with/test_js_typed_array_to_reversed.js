// vybe-test: js/change_array_by_copy_to_reversed_to_spliced_to_sorted_with/test_js_typed_array_to_reversed
// origin: languages/js/tests/js/test_js_change_array_by_copy_to_reversed_to_spliced_to_sorted_with.rs

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

const u8 = new Uint8Array([1, 2, 3]);
const reversed = u8.toReversed();
__check(__line((reversed instanceof Uint8Array) + "|" + reversed.join(",")), "true|3,2,1");
