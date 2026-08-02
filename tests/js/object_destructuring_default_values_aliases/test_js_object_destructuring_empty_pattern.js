// vybe-test: js/object_destructuring_default_values_aliases/test_js_object_destructuring_empty_pattern
// origin: languages/js/tests/js/test_js_object_destructuring_default_values_aliases.rs

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

const {} = { x: 1 };
__check(__line("Empty Pattern Succeeded"), "Empty Pattern Succeeded");
