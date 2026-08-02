// vybe-test: js/ecma/test_setter_auto_invoke
// origin: languages/js/tests/js/js_ecma_test.rs

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
        let t = new Temperature(100);
        __check(__line(t.fahrenheit), "212");
