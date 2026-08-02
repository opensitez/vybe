// vybe-test: js/class_override_getter_setter_super/test_js_class_override_getter_and_setter_both
// origin: languages/js/tests/js/test_js_class_override_getter_setter_super.rs

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
    _score = 0;
    get score() { return this._score; }
    set score(v) { this._score = v; }
}
class Derived extends Base {
    get score() { return super.score + 100; }
    set score(v) { super.score = v * 2; }
}
const d = new Derived();
d.score = 10;
__check(__line(d.score), "120"); // (10 * 2) + 100 = 120
