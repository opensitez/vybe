// vybe-test: js/class_private_fields_get_set_access/test_js_class_private_field_access_on_wrong_object_typeerror
// origin: languages/js/tests/js/test_js_class_private_fields_get_set_access.rs

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
    getBalance(other) {
        return other.#balance; // Accessing #balance on non-BankAccount throws TypeError!
    }
}
const acc = new BankAccount();
try {
    acc.getBalance({});
} catch (e) {
    __check(__line("Private Access Wrong Receiver TypeError"), "Private Access Wrong Receiver TypeError");
}
