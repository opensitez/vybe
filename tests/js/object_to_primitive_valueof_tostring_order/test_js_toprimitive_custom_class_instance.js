// vybe-test: js/object_to_primitive_valueof_tostring_order/test_js_toprimitive_custom_class_instance
// origin: languages/js/tests/js/test_js_object_to_primitive_valueof_tostring_order.rs

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
        if (hint === "number") return this.amount;
        if (hint === "string") return `${this.amount} ${this.currency}`;
        return this.amount;
    }
}
const m = new Money(50, "USD");
__check(__line(`${+m} | ${String(m)} | ${m + 10}`), "50 | 50 USD | 60");
