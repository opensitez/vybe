// vybe-test: js/interop/test_d40_setter_dispatch
// origin: languages/js/tests/js/js_interop_test.rs

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

class Box {
            constructor() { this._value = 0; }
            get value() { return this._value; }
            set value(v) { this._value = v * 2; }
        }
        let b = new Box();
        b.value = 5;
        __check(__line(b.value), "10");
