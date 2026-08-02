// vybe-test: js/prototype_chain_deep/mixin_copies_methods_to_prototype
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

const Serializable = {
    serialize() { return JSON.stringify(this); }
};
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
Object.assign(Point.prototype, Serializable);
const p = new Point(1, 2);
const s = p.serialize();
const parsed = JSON.parse(s);
__check(__line(parsed.x), "1");
__check(__line(parsed.y), "2");
