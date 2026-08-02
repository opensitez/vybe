// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_construct_array_like_arguments
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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

class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
const args = { 0: 5, 1: 15, length: 2 };
const pt = Reflect.construct(Point, args);
__check(__line(`${pt.x},${pt.y}`), "5,15");
