// vybe-test: js/json_serialization/json_custom_toJSON
// origin: languages/js/tests/js/test_json_serialization.rs

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
    constructor(amount, currency) { this.amount = amount; this.currency = currency; }
    toJSON() { return `${this.currency}${this.amount}`; }
}
const obj = { price: new Money(100, "$"), tax: new Money(10, "$") };
const json = JSON.parse(JSON.stringify(obj));
__check(__line(json.price), "$100");
__check(__line(json.tax), "$10");
