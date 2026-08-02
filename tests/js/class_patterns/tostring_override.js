// vybe-test: js/class_patterns/tostring_override
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

class Money {
    constructor(amount, currency) {
        this.amount = amount;
        this.currency = currency;
    }
    toString() { return this.amount + " " + this.currency; }
}
let m = new Money(100, "USD");
__check(__line("" + m), "100 USD");
__check(__line(`${m}`), "100 USD");
