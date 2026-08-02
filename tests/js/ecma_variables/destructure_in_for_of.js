// vybe-test: js/ecma_variables/destructure_in_for_of
// origin: languages/js/tests/js/test_ecma_variables.rs

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

const pairs = [[1, "a"], [2, "b"], [3, "c"]];
for (const [num, letter] of pairs) {
    console.log(num + letter);
}
