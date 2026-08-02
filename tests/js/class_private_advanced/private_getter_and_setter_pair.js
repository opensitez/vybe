// vybe-test: js/class_private_advanced/private_getter_and_setter_pair
// origin: languages/js/tests/js/test_class_private_advanced.rs

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
    get #fahrenheit() { return this.#celsius * 9 / 5 + 32; }
    set #fahrenheit(f) { this.#celsius = (f - 32) * 5 / 9; }
    setF(f) { this.#fahrenheit = f; }
    getC() { return this.#celsius; }
    getF() { return this.#fahrenheit; }
}
const t = new Temperature();
t.setF(212);
__check(__line(t.getC()), "100");
__check(__line(t.getF()), "212");
