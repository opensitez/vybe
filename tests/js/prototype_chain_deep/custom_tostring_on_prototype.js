// vybe-test: js/prototype_chain_deep/custom_tostring_on_prototype
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

function Vector(x, y) { this.x = x; this.y = y; }
Vector.prototype.toString = function() {
    return "(" + this.x + "," + this.y + ")";
};
const v = new Vector(3, 4);
__check(__line(String(v)), "(3,4)");
