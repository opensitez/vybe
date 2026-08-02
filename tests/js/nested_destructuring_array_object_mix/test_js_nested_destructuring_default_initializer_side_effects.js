// vybe-test: js/nested_destructuring_array_object_mix/test_js_nested_destructuring_default_initializer_side_effects
// origin: languages/js/tests/js/test_js_nested_destructuring_array_object_mix.rs

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

let evaluated = false;
const { a: { b = (evaluated = true) } = {} } = { a: { b: "Existing" } };
__check(__line(b + "|Evaluated=" + evaluated), "Existing|Evaluated=false");
