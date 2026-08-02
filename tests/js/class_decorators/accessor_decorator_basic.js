// vybe-test: js/class_decorators/accessor_decorator_basic
// origin: languages/js/tests/js/test_class_decorators.rs

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

class Slider {
    constructor() { this._value = 0; }
    get value() { return this._value; }
    set value(v) { this._value = Math.min(100, Math.max(0, v)); }
}
const s = new Slider();
s.value = 150;
__check(__line(s.value), "100");
s.value = -10;
__check(__line(s.value), "0");
