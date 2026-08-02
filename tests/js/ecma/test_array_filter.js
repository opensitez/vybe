// vybe-test: js/ecma/test_array_filter
// origin: languages/js/tests/js/js_ecma_test.rs

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

let r = [1, 2, 3, 4, 5].filter((x) => x > 3); __check(__line(r.join(",")), "4,5")
