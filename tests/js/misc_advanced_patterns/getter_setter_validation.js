// vybe-test: js/misc_advanced_patterns/getter_setter_validation
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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
    set celsius(v) {
        if (v < -273.15) throw new RangeError("Below absolute zero");
        this.#celsius = v;
    }
    get fahrenheit() { return this.#celsius * 9/5 + 32; }
    set fahrenheit(v) { this.celsius = (v - 32) * 5/9; }
}
const t = new Temperature();
t.celsius = 100;
__check(__line(t.fahrenheit), "212");
t.fahrenheit = 32;
__check(__line(t.celsius), "0");
let threw = false;
try { t.celsius = -300; } catch { threw = true; }
__check(__line(threw), "true");
