// vybe-test: js/typed_arrays/typedarray_bytes_per_element_property
// origin: languages/js/tests/js/test_typed_arrays.rs

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

__check(__line(Int8Array.BYTES_PER_ELEMENT), "1");
__check(__line(Int16Array.BYTES_PER_ELEMENT), "2");
__check(__line(Int32Array.BYTES_PER_ELEMENT), "4");
__check(__line(Float64Array.BYTES_PER_ELEMENT), "8");
