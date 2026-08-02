// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_from_set_iterable
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

const set = new Set([5, 10, 15]);
const i16 = Int16Array.from(set);
__check(__line(i16.join(",")), "5,10,15");
