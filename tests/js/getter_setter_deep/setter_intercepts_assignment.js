// vybe-test: js/getter_setter_deep/setter_intercepts_assignment
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

const obj = {
    _x: 0,
    get x() { return this._x; },
    set x(v) { this._x = v < 0 ? 0 : v; }
};
obj.x = 5;
__check(__line(obj.x), "5");
obj.x = -3;
__check(__line(obj.x), "0");
