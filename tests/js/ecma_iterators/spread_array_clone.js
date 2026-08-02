// vybe-test: js/ecma_iterators/spread_array_clone
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

const orig = [1, 2, 3];
const clone = [...orig];
clone.push(4);
__check(__line(orig.length), "3");
__check(__line(clone.length), "4");
