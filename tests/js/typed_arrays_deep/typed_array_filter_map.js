// vybe-test: js/typed_arrays_deep/typed_array_filter_map
// origin: languages/js/tests/js/test_typed_arrays_deep.rs

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

const arr = new Int32Array([1, 2, 3, 4, 5, 6]);
const evens = arr.filter(x => x % 2 === 0);
const doubled = arr.map(x => x * 2);
__check(__line(evens.join(",")), "2,4,6");
__check(__line(doubled.join(",")), "2,4,6,8,10,12");
