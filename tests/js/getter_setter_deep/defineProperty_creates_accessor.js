// vybe-test: js/getter_setter_deep/defineProperty_creates_accessor
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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
    set(v) { if (Number.isInteger(v)) this._n = v; },
    enumerable: true,
    configurable: true
});
obj.n = 5;
__check(__line(obj.n), "5");
obj.n = 2.5; // ignored — not integer
__check(__line(obj.n), "5");
