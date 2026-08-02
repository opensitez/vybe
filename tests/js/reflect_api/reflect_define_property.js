// vybe-test: js/reflect_api/reflect_define_property
// origin: languages/js/tests/js/test_reflect_api.rs

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

const obj = {};
Reflect.defineProperty(obj, "x", { value: 42, writable: false, enumerable: true, configurable: false });
__check(__line(obj.x), "42");
obj.x = 99; // silently fails — not writable
__check(__line(obj.x), "42");
