// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_sparse_array_source
// origin: languages/js/tests/js/test_js_object_assign_shallow_copy_accessors.rs

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

const sparse = [1, , 3];
const target = Object.assign({}, sparse);
__check(__line(`0=${target[0]}|has1=${1 in target}|2=${target[2]}`), "0=1|has1=false|2=3");
