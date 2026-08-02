// vybe-test: js/ecma_iterators/spread_array_clone_is_shallow
// origin: languages/js/tests/js/test_ecma_iterators.rs

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

const original = [{ x: 1 }];
const copy = [...original];
copy[0].x = 9;
__check(__line(original[0].x), "9");
__check(__line(copy.length), "1");
