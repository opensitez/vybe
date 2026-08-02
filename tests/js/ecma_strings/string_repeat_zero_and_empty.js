// vybe-test: js/ecma_strings/string_repeat_zero_and_empty
// origin: languages/js/tests/js/test_ecma_strings.rs

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

console.log("x".repeat(0) === "");
console.log("".repeat(3));
try {
  console.log("x".repeat(-1));
} catch (e) {
  console.log("RangeError");
}
