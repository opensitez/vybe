// vybe-test: js/array_holes_sparse/map_skips_holes
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

const arr = [1, , 3];
const result = arr.map(x => x * 2);
__check(__line(result[0]), "2");
__check(__line(1 in result), "false");  // hole preserved in map
__check(__line(result[2]), "6");
