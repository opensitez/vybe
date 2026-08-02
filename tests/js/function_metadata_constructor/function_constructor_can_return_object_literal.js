// vybe-test: js/function_metadata_constructor/function_constructor_can_return_object_literal
// origin: languages/js/tests/js/test_function_metadata_constructor.rs

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

const make = new Function("return { x: 1, y: 2 };");
const value = make();
__check(__line(value.x + value.y), "3");
