// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_try_catch_block_body
// origin: languages/js/tests/js/test_js_async_arrow_functions_lexical_this.rs

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

const safeDivide = async (a, b) => {
    try {
        if (b === 0) throw new Error("ZeroDivision");
        return a / b;
    } catch (e) {
        return e.message;
    }
};
safeDivide(10, 0).then(res => console.log(res));
