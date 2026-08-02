// vybe-test: js/misc_advanced_patterns/property_access_order
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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
Object.defineProperty(obj, "computed", {
    get() { return this._x * 2; },
    configurable: true
});
obj._x = 5;
__check(__line(obj.computed), "10");
Object.defineProperty(obj, "computed", {
    get() { return this._x * 3; }
});
__check(__line(obj.computed), "15");
