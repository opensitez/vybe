// vybe-test: js/ecma_classes/class_setter
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

class Temperature {
    constructor(celsius) {
        this._celsius = celsius;
    }
    get fahrenheit() {
        return this._celsius * 9 / 5 + 32;
    }
    set fahrenheit(f) {
        this._celsius = (f - 32) * 5 / 9;
    }
}
const t = new Temperature(0);
__check(__line(t.fahrenheit), "32");
t.fahrenheit = 212;
__check(__line(t._celsius), "100");
