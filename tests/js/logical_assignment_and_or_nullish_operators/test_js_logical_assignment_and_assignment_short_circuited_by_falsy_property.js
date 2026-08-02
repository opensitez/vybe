// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_and_assignment_short_circuited_by_falsy_property
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

let setterCalls = 0;
let rhsExecuted = false;

const obj = {
    _x: 0,
    get x() {
        return this._x;
    },
    set x(v) {
        setterCalls++;
        this._x = v;
    },
};

    obj.x &&= (rhsExecuted = true, 99);
__check(__line(`${obj.x}|${setterCalls}|${rhsExecuted}`), "0|1|false");

obj._x = 1;
rhsExecuted = false;
obj.x &&= (rhsExecuted = true, 33);
__check(__line(`${obj.x}|${setterCalls}|${rhsExecuted}`), "33|2|true");
