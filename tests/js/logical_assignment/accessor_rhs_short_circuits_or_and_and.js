// vybe-test: js/logical_assignment/accessor_rhs_short_circuits_or_and_and
// origin: languages/js/tests/js/test_logical_assignment.rs

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

let calls = 0;
const obj = {
    _x: null,
    get x() {
        calls += 1;
        return this._x;
    },
    set x(v) {
        calls += 1;
        this._x = v;
    },
};

obj.x ||= 10;
obj.x ||= 20;
obj.x &&= 30;

__check(__line(obj._x), "30");
__check(__line(calls), "6");
