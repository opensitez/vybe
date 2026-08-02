// vybe-test: js/ecma/test_private_field
// origin: languages/js/tests/js/js_ecma_test.rs

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
            #balance = 0;
            constructor(initial) {
                this.#balance = initial;
            }
            deposit(amount) {
                this.#balance = this.#balance + amount;
            }
            getBalance() {
                return this.#balance;
            }
        }
        let acc = new BankAccount(100);
        acc.deposit(50);
        __check(__line(acc.getBalance()), "150");
