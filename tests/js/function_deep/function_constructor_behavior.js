// vybe-test: js/function_deep/function_constructor_behavior
// origin: languages/js/tests/js/test_function_deep.rs

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
    this.dist = function() { return Math.sqrt(x*x + y*y); };
}
const p = new Point(3, 4);
__check(__line(p.x), "3");
__check(__line(p.dist()), "5");
__check(__line(p instanceof Point), "true");
