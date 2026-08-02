// vybe-test: js/object_methods_deep/set_prototype_of_rejects_primitive_target
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

let threw = false;
let result;
try {
    result = Object.setPrototypeOf(5, {});
} catch {
    threw = true;
}
__check(__line(threw), "false");
__check(__line(typeof result), "boolean");
__check(__line(result), "false");
