// vybe-test: js/ecma_iterators/spread_object_clone_is_shallow
// origin: languages/js/tests/js/test_ecma_iterators.rs

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

const original = { nested: { x: 1 } };
const copy = { ...original };
copy.nested.x = 5;
__check(__line(original.nested.x), "5");
__check(__line(copy.nested.x), "5");
