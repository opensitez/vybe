// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_custom_properties_on_typedarray_ignored
// origin: languages/js/tests/js/test_js_structured_clone_typed_arrays_array_buffers.rs

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

const u8 = new Uint8Array([1, 2]);
u8.customMeta = "metadata";
const clone = structuredClone(u8);
__check(__line(clone.length + "|hasMeta=" + Object.hasOwn(clone, "customMeta")), "2|hasMeta=false");
