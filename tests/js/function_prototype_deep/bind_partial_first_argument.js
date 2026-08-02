// vybe-test: js/function_prototype_deep/bind_partial_first_argument
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

function sub(a, b) { return a - b; } const fromTen = sub.bind(null, 10); __check(__line(fromTen(3)), "7");
