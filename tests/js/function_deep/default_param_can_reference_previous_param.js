// vybe-test: js/function_deep/default_param_can_reference_previous_param
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

function f(x, y = x * 2, z = x + y) {
    return `${x},${y},${z}`;
}
__check(__line(f(3)), "3,6,9");
__check(__line(f(3, 10)), "3,10,13");
