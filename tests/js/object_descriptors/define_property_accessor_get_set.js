// vybe-test: js/object_descriptors/define_property_accessor_get_set
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const obj = { _n: 0 };
Object.defineProperty(obj, "n", {
    get() { return this._n; },
    set(v) { this._n = v < 0 ? 0 : v; },
    configurable: true
});
obj.n = 5;
__check(__line(obj.n), "5");
obj.n = -3;
__check(__line(obj.n), "0");
