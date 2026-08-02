// vybe-test: js/property_descriptor_deep/define_accessor_property
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

const obj = { _x: 0 };
Object.defineProperty(obj, "x", {
    get() { return this._x; },
    set(v) { this._x = v * 2; },
    configurable: true, enumerable: true
});
obj.x = 5;
__check(__line(obj.x), "10");
