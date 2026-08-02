// vybe-test: js/array_holes_sparse/copywithin_preserves_length_and_holes
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

const arr = [1, , 2, , 3];
const copied = arr.copyWithin(1, 0, 2);
__check(__line(copied.length), "5");
__check(__line(2 in copied), "true");
__check(__line(copied.join(",")), "1,,,,3");
