// vybe-test: js/object_descriptors/object_spread_shallow_copy
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const src = { a: 1, b: { nested: 2 } };
const copy = { ...src };
copy.a = 99;
copy.b.nested = 88;
__check(__line(src.a), "1");
__check(__line(src.b.nested), "88");
