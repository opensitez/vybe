// vybe-test: js/array_map_filter_reduce_reduce_right/test_js_array_map_this_arg_binding
// origin: languages/js/tests/js/test_js_array_map_filter_reduce_reduce_right.rs

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

const context = { multiplier: 10 };
const nums = [1, 2];
const res = nums.map(function(x) { return x * this.multiplier; }, context);
__check(__line(res.join(",")), "10,20");
