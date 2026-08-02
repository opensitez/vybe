// vybe-test: js/arrow_function_lexical_this_arguments_super/test_js_arrow_function_in_array_higher_order_methods
// origin: languages/js/tests/js/test_js_arrow_function_lexical_this_arguments_super.rs

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

const arr = [1, 2, 3];
const doubled = arr.map(x => x * 2);
__check(__line(doubled.join(",")), "2,4,6");
