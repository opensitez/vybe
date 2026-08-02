// vybe-test: js/function_bind_currying_bound_this/test_js_bound_function_constructor_behavior_ignores_bound_this
// origin: languages/js/tests/js/test_js_function_bind_currying_bound_this.rs

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

function Point(x, y) {
    this.x = x;
    this.y = y;
}
const BoundPoint = Point.bind({ x: 99, y: 99 }, 10);
const p = new BoundPoint(20); // new operator ignores bound 'this' context, but retains prepended arguments!
__check(__line(`${p.x}:${p.y}`), "10:20");
