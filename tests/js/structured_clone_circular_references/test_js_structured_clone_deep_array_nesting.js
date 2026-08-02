// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_deep_array_nesting
// origin: languages/js/tests/js/test_js_structured_clone_circular_references.rs

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

const nested = [[["deep"]]];
const clone = structuredClone(nested);
__check(__line(clone[0][0][0]), "deep");
