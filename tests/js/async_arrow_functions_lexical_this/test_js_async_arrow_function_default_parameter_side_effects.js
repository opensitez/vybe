// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_default_parameter_side_effects
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

let count = 0;
const getDefault = () => ++count;
const fn = async (val = getDefault()) => val * 10;

fn().then(r1 => {
    fn(100).then(r2 => {
        console.log(`${r1},${r2}|count=${count}`);
    });
});
