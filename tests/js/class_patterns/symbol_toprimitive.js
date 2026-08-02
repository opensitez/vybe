// vybe-test: js/class_patterns/symbol_toprimitive
// origin: languages/js/tests/js/test_class_patterns.rs

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
    constructor(celsius) { this.celsius = celsius; }
    [Symbol.toPrimitive](hint) {
        if (hint === "number") return this.celsius;
        if (hint === "string") return this.celsius + "°C";
        return this.celsius;
    }
}
let t = new Temperature(100);
__check(__line(+t), "100");
__check(__line(`${t}`), "100°C");
