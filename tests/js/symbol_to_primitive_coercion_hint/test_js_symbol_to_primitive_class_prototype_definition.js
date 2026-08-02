// vybe-test: js/symbol_to_primitive_coercion_hint/test_js_symbol_to_primitive_class_prototype_definition
// origin: languages/js/tests/js/test_js_symbol_to_primitive_coercion_hint.rs

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

class Money {
    constructor(amount, currency) {
        this.amount = amount;
        this.currency = currency;
    }
    [Symbol.toPrimitive](hint) {
        if (hint === "string") return `${this.amount} ${this.currency}`;
        return this.amount;
    }
}
const m = new Money(50, "USD");
__check(__line(String(m) + "|" + (m + 10)), "50 USD|60");
