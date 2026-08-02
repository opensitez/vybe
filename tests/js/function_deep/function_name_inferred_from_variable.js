// vybe-test: js/function_deep/function_name_inferred_from_variable
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

const myFunc = function() {};
const arrow = () => {};
__check(__line(myFunc.name), "myFunc");
__check(__line(arrow.name), "arrow");
