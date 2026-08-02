// vybe-test: js/module_patterns/revealing_module_pattern
// origin: languages/js/tests/js/test_module_patterns.rs

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

const BankAccount = (() => {
    let balance = 0;
    const deposit = (n) => { balance += n; };
    const withdraw = (n) => { if (n > balance) throw new Error("insufficient"); balance -= n; };
    const getBalance = () => balance;
    return { deposit, withdraw, getBalance };
})();
BankAccount.deposit(100);
BankAccount.deposit(50);
BankAccount.withdraw(30);
__check(__line(BankAccount.getBalance()), "120");
