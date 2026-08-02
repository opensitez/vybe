// vybe-test: js/inheritance/test_18_setter_on_base
// origin: languages/js/tests/js/js_inheritance_test.rs

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

class Container {
            constructor() { this._data = 0; }
            get data() { return this._data; }
            set data(v) { this._data = v + 1; }
        }
        let c = new Container();
        c.data = 9;
        __check(__line(c.data), "10");
