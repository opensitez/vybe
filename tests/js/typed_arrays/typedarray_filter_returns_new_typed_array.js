// vybe-test: js/typed_arrays/typedarray_filter_returns_new_typed_array
// origin: languages/js/tests/js/test_typed_arrays.rs

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

const a = new Int32Array([1, 2, 3, 4, 5, 6]);
const evens = a.filter(x => x % 2 === 0);
__check(__line(evens.length), "3");
__check(__line(evens[0]), "2");
__check(__line(evens[1]), "4");
__check(__line(evens[2]), "6");
