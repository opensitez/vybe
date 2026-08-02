// vybe-test: js/ecma_modules/export_function
// origin: languages/js/tests/js/test_ecma_modules.rs

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

export function add(a, b) {
    return a + b;
}
__check(__line(add(1, 2)), "3");
