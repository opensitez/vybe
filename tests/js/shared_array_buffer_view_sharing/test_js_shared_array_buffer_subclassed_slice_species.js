// vybe-test: js/shared_array_buffer_view_sharing/test_js_shared_array_buffer_subclassed_slice_species
// origin: languages/js/tests/js/test_js_shared_array_buffer_view_sharing.rs

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

class CustomSAB extends SharedArrayBuffer {}
const csab = new CustomSAB(8);
const sliced = csab.slice(0, 4);
__check(__line(sliced instanceof SharedArrayBuffer), "true");
