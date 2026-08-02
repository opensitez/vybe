// vybe-test: js/oop_patterns_advanced/lazy_initialization_class
// origin: languages/js/tests/js/test_oop_patterns_advanced.rs

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

class ExpensiveResource {
    #_data = null;
    get data() {
        if (!this.#_data) this.#_data = { computed: 42 };
        return this.#_data;
    }
}
const r = new ExpensiveResource();
__check(__line(r.data.computed), "42");
__check(__line(r.data === r.data), "true");
