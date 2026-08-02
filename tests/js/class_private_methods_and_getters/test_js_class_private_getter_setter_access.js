// vybe-test: js/class_private_methods_and_getters/test_js_class_private_getter_setter_access
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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
    get #fahrenheit() { return this.#celsius * 1.8 + 32; }
    set #fahrenheit(f) { this.#celsius = (f - 32) / 1.8; }

    setFahrenheit(f) { this.#fahrenheit = f; }
    getFahrenheit() { return this.#fahrenheit; }
}
const t = new Temperature();
t.setFahrenheit(100);
__check(__line(t.getFahrenheit().toFixed(1)), "100.0");
