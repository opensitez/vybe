// vybe-test: js/array_hole_length_basics/sparse_array_map_preserves_hole_positions
// origin: languages/js/tests/js/test_array_hole_length_basics.rs

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

const arr = [, 2, , 4];
const mapped = arr.map(x => x * 2);
__check(__line(mapped.length), "4");
__check(__line(0 in mapped), "false");
__check(__line(1 in mapped), "true");
__check(__line(2 in mapped), "false");
__check(__line(3 in mapped), "true");
__check(__line(mapped[1]), "4");
