// vybe-test: js/array_holes_sparse/spread_fills_holes_with_undefined
// origin: languages/js/tests/js/test_array_holes_sparse.rs

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

const sparse = [1, , 3];
const dense = [...sparse];
__check(__line(1 in dense), "true"); // true — undefined, not hole
__check(__line(dense[1]), "undefined");
