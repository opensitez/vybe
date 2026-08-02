// vybe-test: js/function_arguments_hoisting/function_name_property
// origin: languages/js/tests/js/test_function_arguments_hoisting.rs

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

function named() {}
const anon = function() {};
const arrow = () => {};
__check(__line(named.name), "named");
__check(__line(anon.name), "anon");
__check(__line(arrow.name), "arrow");
