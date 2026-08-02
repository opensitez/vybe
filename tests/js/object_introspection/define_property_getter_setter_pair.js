// vybe-test: js/object_introspection/define_property_getter_setter_pair
// origin: languages/js/tests/js/test_object_introspection.rs

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

const obj = { _count: 0 };
Object.defineProperty(obj, "count", {
    get() { return this._count; },
    set(v) { this._count = v < 0 ? 0 : v; },
    enumerable: true,
    configurable: true
});
obj.count = 5;
__check(__line(obj.count), "5");
obj.count = -3;
__check(__line(obj.count), "0");
