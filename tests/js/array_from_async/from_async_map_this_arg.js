// vybe-test: js/array_from_async/from_async_map_this_arg
// origin: languages/js/tests/js/test_array_from_async.rs

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

const ctx = { multiplier: 5 };
Array.fromAsync([1, 2, 3], function(x) {
    return x * this.multiplier;
}, ctx).then(arr => console.log(arr.join(",")));
