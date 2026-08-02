// vybe-test: js/object_introspection/tostring_override_via_prototype
// origin: languages/js/tests/js/test_object_introspection.rs

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
Point.prototype.toString = function() {
    return "(" + this.x + "," + this.y + ")";
};
const p = new Point(3, 4);
__check(__line(p.toString()), "(3,4)");
__check(__line(String(p)), "(3,4)");
