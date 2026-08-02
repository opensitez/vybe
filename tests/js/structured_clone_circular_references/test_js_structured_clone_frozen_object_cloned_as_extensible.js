// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_frozen_object_cloned_as_extensible
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

const frozen = Object.freeze({ a: 1 });
const clone = structuredClone(frozen);
__check(__line(Object.isFrozen(clone) + "|" + Object.isExtensible(clone)), "false|true");
