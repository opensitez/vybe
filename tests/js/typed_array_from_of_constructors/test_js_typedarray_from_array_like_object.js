// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_from_array_like_object
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

const arrayLike = { 0: 10, 1: 20, length: 2 };
const u8 = Uint8Array.from(arrayLike);
__check(__line(u8.join(",")), "10,20");
