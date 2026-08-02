// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_instanceof_symbol_hasinstance_observes_argument
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

const log = [];
class Tracker {
    static [Symbol.hasInstance](value) {
        log.push(typeof value);
        return value && value.token === "ok";
    }
}
__check(__line({ token: "ok" } instanceof Tracker), "true");
__check(__line({ token: "bad" } instanceof Tracker), "false");
__check(__line(log.join("|")), "object|object");
