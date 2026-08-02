// vybe-test: js/class_private_advanced/private_field_in_static_factory_method
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

class Color {
    #r; #g; #b;
    constructor(r, g, b) { this.#r = r; this.#g = g; this.#b = b; }
    static fromHex(hex) {
        const r = parseInt(hex.slice(1, 3), 16);
        const g = parseInt(hex.slice(3, 5), 16);
        const b = parseInt(hex.slice(5, 7), 16);
        return new Color(r, g, b);
    }
    toString() { return this.#r + "," + this.#g + "," + this.#b; }
}
const c = Color.fromHex("#ff8000");
__check(__line(c.toString()), "255,128,0");
