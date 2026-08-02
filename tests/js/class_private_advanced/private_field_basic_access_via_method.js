// vybe-test: js/class_private_advanced/private_field_basic_access_via_method
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

class BankAccount {
    #balance = 100;
    getBalance() { return this.#balance; }
}
const acc = new BankAccount();
__check(__line(acc.getBalance()), "100");
