// vybe-test: js/object_immutability/seal_keeps_object_sealed_and_blocks_extension
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

const obj = Object.seal({ x: 1 });
obj.x = 99;
obj.y = 2;
__check(__line(obj.x), "99");
__check(__line(Object.prototype.hasOwnProperty.call(obj, "y")), "false");
__check(__line(Object.isSealed(obj)), "true");
