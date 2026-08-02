// vybe-test: js/array_es2023/array_spread_shallow_copy
// origin: languages/js/tests/js/test_array_es2023.rs

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

const orig = [1, 2, 3];
const copy = [...orig];
copy.push(4);
__check(__line(orig.join(",")), "1,2,3");
__check(__line(copy.join(",")), "1,2,3,4");
__check(__line(orig === copy), "false");
