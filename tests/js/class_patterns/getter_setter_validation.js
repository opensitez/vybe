// vybe-test: js/class_patterns/getter_setter_validation
// origin: languages/js/tests/js/test_class_patterns.rs

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

class User {
    #email;
    get email() { return this.#email; }
    set email(val) {
        if (!val.includes("@")) throw new Error("invalid email");
        this.#email = val;
    }
}
let u = new User();
u.email = "alice@test.com";
__check(__line(u.email), "alice@test.com");
try {
    u.email = "invalid";
} catch (e) {
    __check(__line(e.message), "invalid email");
}
