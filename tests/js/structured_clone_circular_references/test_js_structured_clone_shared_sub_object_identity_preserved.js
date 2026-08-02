// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_shared_sub_object_identity_preserved
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

const shared = { val: 42 };
const root = { first: shared, second: shared };
const clone = structuredClone(root);
__check(__line((clone.first !== shared) + "|" + (clone.first === clone.second)), "true|true");
