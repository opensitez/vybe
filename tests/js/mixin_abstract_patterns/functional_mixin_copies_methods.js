// vybe-test: js/mixin_abstract_patterns/functional_mixin_copies_methods
// origin: languages/js/tests/js/test_mixin_abstract_patterns.rs

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

function applyMixin(target, mixin) {
    Object.assign(target.prototype, mixin);
}
const Serializable = {
    serialize() { return JSON.stringify(this); }
};
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
applyMixin(Point, Serializable);
const p = new Point(1, 2);
__check(__line(p.serialize()), "{\"x\":1,\"y\":2}");
