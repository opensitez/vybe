// vybe-test: js/array_methods_new/flat_infinity_depth
// origin: languages/js/tests/js/test_array_methods_new.rs

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

const deeply = [[[[[1], 2], 3], 4], 5];
__check(__line(deeply.flat(Infinity).join(",")), "1,2,3,4,5");
