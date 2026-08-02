// vybe-test: js/features/test_while_loop
// origin: languages/js/tests/js/js_features_test.rs

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

let i = 0; while (i < 5) { i = i + 1; } console.log(i)
