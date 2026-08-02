// vybe-test: js/logical_assignment/and_assign_preserves_falsy_without_setter
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
    _x: "",
    get x() {
        calls += 1;
        return this._x;
    },
    set x(v) {
        calls += 1;
        this._x = v;
    },
};

obj.x &&= 99;
obj._x = "value";
obj.x &&= 77;

__check(__line(obj._x), "77");
__check(__line(calls), "4");
