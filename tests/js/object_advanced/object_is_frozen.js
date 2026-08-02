// vybe-test: js/object_advanced/object_is_frozen
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

let obj = { a: 1 };
__check(__line(Object.isFrozen(obj)), "false");
Object.freeze(obj);
__check(__line(Object.isFrozen(obj)), "true");
