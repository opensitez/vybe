// vybe-test: js/class_private_deep/private_accessor_auto_accessor
// origin: languages/js/tests/js/test_class_private_deep.rs

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
    #celsius = 0;
    get celsius() { return this.#celsius; }
    set celsius(v) { this.#celsius = v; }
    get fahrenheit() { return this.#celsius * 9/5 + 32; }
}
const t = new Temperature();
t.celsius = 100;
__check(__line(t.celsius), "100");
__check(__line(t.fahrenheit), "212");
