// vybe-test: js/array_hole_length_basics/array_constructor_with_negative_length_throws_range_error
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

try {
  new Array(-1);
  console.log("no error");
} catch (error) {
  console.log(error instanceof RangeError);
}
