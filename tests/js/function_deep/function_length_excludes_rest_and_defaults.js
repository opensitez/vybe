// vybe-test: js/function_deep/function_length_excludes_rest_and_defaults
// origin: languages/js/tests/js/test_function_deep.rs

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

function f1(a, b, c) {}
function f2(a, b = 1, c) {} // default stops counting
function f3(a, ...rest) {}
__check(__line(f1.length), "3");
__check(__line(f2.length), "1");
__check(__line(f3.length), "1");
