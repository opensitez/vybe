// vybe-test: js/structured_clone_typed_arrays_array_buffers/test_js_structured_clone_typedarray_inside_object_graph
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

const data = {
    name: "Dataset",
    values: new Float32Array([10.5, 20.5])
};
const clone = structuredClone(data);
__check(__line(clone.name + "|" + (clone.values instanceof Float32Array) + "|" + clone.values.join(",")), "Dataset|true|10.5,20.5");
