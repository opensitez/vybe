// vybe-test: js/scope_prototype/class_private_method
// origin: languages/js/tests/js/test_scope_prototype.rs

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

class Processor {
    #transform(x) { return x * 2; }
    process(x) { return this.#transform(x) + 1; }
}
let p = new Processor();
__check(__line(p.process(5)), "11");
