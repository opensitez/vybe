// vybe-test: js/scope_closures_lexical_environment_capture/test_js_closure_private_data_hiding_pattern
// origin: languages/js/tests/js/test_js_scope_closures_lexical_environment_capture.rs

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

function BankAccount(initialBalance) {
    let balance = initialBalance;
    this.deposit = (amt) => { balance += amt; };
    this.getBalance = () => balance;
}
const acc = new BankAccount(100);
acc.deposit(50);
__check(__line(acc.getBalance() + "|hasBalanceProp=" + ("balance" in acc)), "150|hasBalanceProp=false");
