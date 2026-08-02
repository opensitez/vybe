// vybe-test: js/array_find_last_and_find_last_index/test_js_array_find_last_this_arg_binding
// origin: languages/js/tests/js/test_js_array_find_last_and_find_last_index.rs

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

const ctx = { threshold: 20 };
const arr = [10, 25, 30];
const found = arr.findLast(function(val) {
    return val > this.threshold;
}, ctx);
__check(__line(found), "30");
