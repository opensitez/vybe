// vybe-test: js/prototype_chain_deep/custom_valueof_used_in_arithmetic
// origin: languages/js/tests/js/test_prototype_chain_deep.rs

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

function Money(amount) { this.amount = amount; }
Money.prototype.valueOf = function() { return this.amount; };
const m1 = new Money(10);
const m2 = new Money(20);
__check(__line(m1 + m2), "30");
