// vybe-test: js/typed_array_advanced/typed_array_subarray_is_view
// origin: languages/js/tests/js/test_typed_array_advanced.rs

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

const ta = new Int32Array([10, 20, 30, 40, 50]);
const sub = ta.subarray(1, 3);
__check(__line(sub.length), "2");
__check(__line(sub[0]), "20");
__check(__line(sub[1]), "30");
// Modifying subarray affects original
sub[0] = 99;
__check(__line(ta[1]), "99");
