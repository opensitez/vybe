// vybe-test: js/array_hole_length_basics/sparse_array_for_each_skips_holes
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

const arr = [, "a", , "b"];
const seen = [];
arr.forEach((value, index) => seen.push(index + ":" + value));
console.log(seen.join(","));
