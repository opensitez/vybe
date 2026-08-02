// vybe-test: js/ecma_operators/accessor_property_arithmetic_assignment_uses_getter_setter
// origin: languages/js/tests/js/test_ecma_operators.rs

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
    _x: 1,
    get x() {
        return this._x;
    },
    set x(v) {
        this._x = v;
    }
};

obj.x += 4;
__check(__line(obj.x), "5");
__check(__line(obj._x), "5");
