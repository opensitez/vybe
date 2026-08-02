// vybe-test: js/generator_function_prototype/generator_symbol_method_name_is_symbol_description
// origin: languages/js/tests/js/test_generator_function_prototype.rs

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

const sym = Symbol("iter"); const obj = { *[sym]() { yield 1; } }; __check(__line(obj[sym].name), "[iter]");
