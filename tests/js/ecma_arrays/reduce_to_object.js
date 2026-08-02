// vybe-test: js/ecma_arrays/reduce_to_object
// origin: languages/js/tests/js/test_ecma_arrays.rs

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

const pairs = [["a", 1], ["b", 2], ["c", 3]];
const obj = pairs.reduce((acc, [k, v]) => {
    acc[k] = v;
    return acc;
}, {});
__check(__line(obj.a), "1");
__check(__line(obj.b), "2");
__check(__line(obj.c), "3");
