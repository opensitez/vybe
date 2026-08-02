// vybe-test: js/class_fields_advanced/auto_accessor_generates_getter_and_setter
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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

class Temp {
    accessor celsius = 0;
    get fahrenheit() { return this.celsius * 1.8 + 32; }
}
const t = new Temp();
t.celsius = 100;
__check(__line(t.fahrenheit), "212");
__check(__line(t.celsius), "100");
