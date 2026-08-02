// vybe-test: js/function_call_apply_bind/bind_then_new_ignores_this
// origin: languages/js/tests/js/test_function_call_apply_bind.rs

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

function Point(x, y) { this.x = x; this.y = y; }
const obj = { name: "ignored" };
const BoundPoint = Point.bind(obj, 1); // bind this and first arg
const p = new BoundPoint(2); // new ignores bound this
__check(__line(p.x), "1");
__check(__line(p.y), "2");
__check(__line(p instanceof Point), "true");
