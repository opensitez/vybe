// vybe-test: js/advanced/test_array_of_functions
// origin: languages/js/tests/js/js_advanced_test.rs

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

let fns = [(x) => x + 1, (x) => x * 2, (x) => x * x];
        __check(__line(fns[0](5), fns[1](5), fns[2](5)), "6 10 25");
