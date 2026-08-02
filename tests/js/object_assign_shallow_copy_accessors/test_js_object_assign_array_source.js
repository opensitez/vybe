// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_array_source
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

const target = Object.assign({}, ["x", "y", "z"]);
__check(__line(`${target[0]}:${target[1]}:${target[2]}`), "x:y:z");
