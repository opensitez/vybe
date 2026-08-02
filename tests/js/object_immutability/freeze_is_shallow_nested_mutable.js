// vybe-test: js/object_immutability/freeze_is_shallow_nested_mutable
// origin: languages/js/tests/js/test_object_immutability.rs

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

const obj = Object.freeze({ nested: { x: 1 } });
obj.nested.x = 99; // nested not frozen
__check(__line(obj.nested.x), "99");
