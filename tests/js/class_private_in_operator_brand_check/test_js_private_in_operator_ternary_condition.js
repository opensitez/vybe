// vybe-test: js/class_private_in_operator_brand_check/test_js_private_in_operator_ternary_condition
// origin: languages/js/tests/js/test_js_class_private_in_operator_brand_check.rs

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
    #balance = 100;
    static read(obj) {
        return #balance in obj ? obj.#balance : -1;
    }
}
const a = new Account();
__check(__line(Account.read(a) + "|" + Account.read({})), "100|-1");
