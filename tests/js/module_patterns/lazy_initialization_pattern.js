// vybe-test: js/module_patterns/lazy_initialization_pattern
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

class LazyInit {
    #instance = null;
    #factory;
    constructor(factory) { this.#factory = factory; }
    get() {
        if (!this.#instance) {
            this.#instance = this.#factory();
        }
        return this.#instance;
    }
}
let created = 0;
const lazy = new LazyInit(() => { created++; return { value: 42 }; });
__check(__line(created), "0");     // not yet created
const v = lazy.get();
__check(__line(created), "1");     // now created
const v2 = lazy.get();
__check(__line(created), "1");     // not created again
__check(__line(v === v2), "true");    // same instance
