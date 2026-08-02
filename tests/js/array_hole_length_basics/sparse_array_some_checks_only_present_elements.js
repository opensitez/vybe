// vybe-test: js/array_hole_length_basics/sparse_array_some_checks_only_present_elements
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

const arr = [, , 3];
const seen = [];
const result = arr.some((value, index) => {
  seen.push(index);
  return value === 3;
});
__check(__line(result), "true");
__check(__line(seen.join(",")), "2");
