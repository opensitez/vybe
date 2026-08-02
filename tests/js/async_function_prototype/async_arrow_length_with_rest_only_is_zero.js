// vybe-test: js/async_function_prototype/async_arrow_length_with_rest_only_is_zero
// origin: languages/js/tests/js/test_async_function_prototype.rs

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

const f = async (...xs) => xs; __check(__line(f.length), "0");
