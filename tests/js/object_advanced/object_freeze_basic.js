// vybe-test: js/object_advanced/object_freeze_basic
// origin: languages/js/tests/js/test_object_advanced.rs

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

let obj = { x: 1, y: 2 };
Object.freeze(obj);
obj.x = 99;
obj.z = 3;
__check(__line(obj.x), "1");
__check(__line(obj.z), "undefined");
