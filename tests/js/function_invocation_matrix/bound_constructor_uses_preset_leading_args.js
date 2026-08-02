// vybe-test: js/function_invocation_matrix/bound_constructor_uses_preset_leading_args
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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
const PointY2 = Point.bind(null, 1);
const p = new PointY2(2);
__check(__line(p.x), "1");
__check(__line(p.y), "2");
