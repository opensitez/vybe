// vybe-test: js/class_private_advanced/private_method_calling_another_private_method
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class StringProcessor {
    #trim(s) { return s.trim(); }
    #upper(s) { return s.toUpperCase(); }
    #process(s) { return this.#upper(this.#trim(s)); }
    run(s) { return this.#process(s); }
}
const sp = new StringProcessor();
__check(__line(sp.run("  hello world  ")), "HELLO WORLD");
