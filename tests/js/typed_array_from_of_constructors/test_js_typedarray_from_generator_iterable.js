// vybe-test: js/typed_array_from_of_constructors/test_js_typedarray_from_generator_iterable
// origin: languages/js/tests/js/test_js_typed_array_from_of_constructors.rs

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

function* gen() { yield 2; yield 4; yield 6; }
const f64 = Float64Array.from(gen());
__check(__line(f64.join(",")), "2,4,6");
