// vybe-test: js/object_assign_shallow_copy_accessors/test_js_object_assign_merging_own_enumerable_properties
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

const target = { a: 1 };
const source1 = { b: 2 };
const source2 = { c: 3 };
const res = Object.assign(target, source1, source2);
__check(__line(`${res.a}:${res.b}:${res.c}:${res === target}`), "1:2:3:true");
