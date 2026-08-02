// vybe-test: js/ecma/test_getter_setter_pair
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

class Account {
            constructor(balance) {
                this._balance = balance;
            }
            get balance() {
                return this._balance;
            }
            set balance(val) {
                if (val < 0) { this._balance = 0; }
                else { this._balance = val; }
            }
        }
        let a = new Account(100);
        __check(__line(a.balance), "100");
        a.balance = 200;
        __check(__line(a.balance), "200");
        a.balance = -50;
        __check(__line(a.balance), "0");
