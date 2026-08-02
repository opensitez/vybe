// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_apply_array_like_arguments
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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

function fn(a, b) { return a + b; }
const args = { 0: 10, 1: 20, length: 2 };
__check(__line(Reflect.apply(fn, null, args)), "30");
