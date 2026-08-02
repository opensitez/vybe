// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_non_enumerable_properties_ignored
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

const obj = { visible: 1 };
Object.defineProperty(obj, "hidden", { value: 2, enumerable: false });
const clone = structuredClone(obj);
__check(__line(clone.visible + "|hasHidden=" + ("hidden" in clone)), "1|hasHidden=false");
