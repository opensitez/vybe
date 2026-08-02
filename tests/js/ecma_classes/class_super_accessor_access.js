// vybe-test: js/ecma_classes/class_super_accessor_access
// origin: languages/js/tests/js/test_ecma_classes.rs

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

class Base {
    get score() { return this._score; }
    set score(v) { this._score = v; }
}
class Derived extends Base {
    set score(v) { super.score = v + 1; }
    get score() { return super.score * 2; }
}
const d = new Derived();
d.score = 3;
__check(__line(d.score), "8");
__check(__line(d._score), "4");
