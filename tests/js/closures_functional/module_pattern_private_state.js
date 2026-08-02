// vybe-test: js/closures_functional/module_pattern_private_state
// origin: languages/js/tests/js/test_closures_functional.rs

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

let bank = (function() {
    let balance = 0;
    return {
        deposit(amt) { balance += amt; },
        withdraw(amt) {
            if (amt > balance) return false;
            balance -= amt;
            return true;
        },
        getBalance() { return balance; }
    };
})();
bank.deposit(100);
bank.deposit(50);
__check(__line(bank.withdraw(30)), "true");
__check(__line(bank.getBalance()), "120");
__check(__line(bank.withdraw(200)), "false");
__check(__line(bank.getBalance()), "120");
