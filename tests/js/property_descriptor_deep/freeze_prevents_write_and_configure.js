// vybe-test: js/property_descriptor_deep/freeze_prevents_write_and_configure
// origin: languages/js/tests/js/test_property_descriptor_deep.rs

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

const obj = { x: 1 };
Object.freeze(obj);
obj.x = 99;
obj.y = 2;
__check(__line(obj.x), "1");
__check(__line(obj.y), "undefined");
__check(__line(Object.isFrozen(obj)), "true");
