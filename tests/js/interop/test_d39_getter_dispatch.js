// vybe-test: js/interop/test_d39_getter_dispatch
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

class Temperature {
            constructor(celsius) { this._c = celsius; }
            get fahrenheit() { return this._c * 9 / 5 + 32; }
        }
        let t = new Temperature(100);
        __check(__line(t.fahrenheit), "212");
