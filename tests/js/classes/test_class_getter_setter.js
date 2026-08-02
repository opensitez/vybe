// vybe-test: js/classes/test_class_getter_setter
// origin: languages/js/tests/js/js_classes_test.rs

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
                this.celsius = celsius;
            }
            get fahrenheit() {
                return this.celsius * 9 / 5 + 32;
            }
            set fahrenheit(value) {
                this.celsius = (value - 32) * 5 / 9;
            }
        }
        const t = new Temperature(0);
        const f = t.fahrenheit;
        t.fahrenheit = 212;
        __check(__line(f, t.celsius), "32 100");
