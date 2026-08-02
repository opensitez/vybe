// vybe-test: js/structured_clone_circular_references/test_js_structured_clone_getter_setter_evaluated_to_data_property
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

const obj = {
    get val() { return 100; }
};
const clone = structuredClone(obj);
const desc = Object.getOwnPropertyDescriptor(clone, "val");
__check(__line(desc.value + "|hasGetter=" + (typeof desc.get !== "undefined")), "100|hasGetter=false");
