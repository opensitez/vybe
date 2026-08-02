// vybe-test: js/unary_plus_minus_tilde_void_typeof_delete/test_js_delete_array_element_creates_hole
// origin: languages/js/tests/js/test_js_unary_plus_minus_tilde_void_typeof_delete.rs

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

const arr = [10, 20, 30];
const res = delete arr[1];
__check(__line(res + "|len=" + arr.length + "|hasHole=" + !(1 in arr)), "true|len=3|hasHole=true");
