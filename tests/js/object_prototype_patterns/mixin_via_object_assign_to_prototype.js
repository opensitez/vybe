// vybe-test: js/object_prototype_patterns/mixin_via_object_assign_to_prototype
// origin: languages/js/tests/js/test_object_prototype_patterns.rs

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
    toJSON() { return JSON.stringify(this); },
    fromJSON(str) { return Object.assign(Object.create(this), JSON.parse(str)); }
};
class Point {
    constructor(x, y) { this.x = x; this.y = y; }
}
Object.assign(Point.prototype, Serializable);
const p = new Point(3, 4);
const json = JSON.stringify(p);
const p2 = JSON.parse(json);
__check(__line(p2.x), "3");
__check(__line(p2.y), "4");
