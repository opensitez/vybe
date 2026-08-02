// vybe-test: js/function_edge_cases/getter_in_object_literal
// origin: languages/js/tests/js/test_function_edge_cases.rs

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
    set x(v) { this._x = v * 2; }
};
obj.x = 5;
__check(__line(obj.x), "10");
