// vybe-test: js/object_prototype_patterns/object_create_with_accessor_in_props
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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

const obj = Object.create({}, {
    x: {
        get() { return this._x ?? 0; },
        set(v) { this._x = v; },
        configurable: true,
        enumerable: true
    }
});
obj.x = 42;
__check(__line(obj.x), "42");
