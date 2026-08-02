// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_primitive_sources_coerced_only_strings_have_enumerable_own
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

const target = Object.assign({}, "abc", 123, true);
__check(__line(Object.values(target).join(",")), "a,b,c"); // String primitives are wrapped, indices 0,1,2 copied!
