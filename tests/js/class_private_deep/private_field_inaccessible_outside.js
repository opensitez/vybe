// vybe-test: js/class_private_deep/private_field_inaccessible_outside
// origin: languages/js/tests/js/test_class_private_deep.rs

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

class Wallet {
    #balance = 0;
    deposit(n) { this.#balance += n; }
    get balance() { return this.#balance; }
}
const w = new Wallet();
w.deposit(100);
__check(__line(w.balance), "100");
const key = "#" + "balance";
__check(__line(w[key] === undefined), "true");
