// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_lexical_this_in_object_method
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

const obj = {
    multiplier: 3,
    calc(numbers) {
        // Async arrow preserves enclosing 'this'
        const fn = async (n) => n * this.multiplier;
        return Promise.all(numbers.map(fn));
    }
};
obj.calc([1, 2, 3]).then(results => console.log(results.join(",")));
