// vybe-test: js/array_hole_length_basics/sparse_array_every_ignores_holes_for_predicate_calls
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

const arr = [, , 4];
const seen = [];
const result = arr.every((value, index) => {
  seen.push(index);
  return value > 0;
});
__check(__line(result), "true");
__check(__line(seen.join(",")), "2");
