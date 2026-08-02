// vybe-test: js/array_from_patterns/array_from_this_arg_context
// origin: languages/js/tests/js/test_array_from_patterns.rs

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

const ctx = { scale: 3 };
const arr = Array.from([1, 2, 3], function(x) {
    return x * this.scale;
}, ctx);
__check(__line(arr.join(",")), "3,6,9");
