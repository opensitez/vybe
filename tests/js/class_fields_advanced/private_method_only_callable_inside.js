// vybe-test: js/class_fields_advanced/private_method_only_callable_inside
// origin: languages/js/tests/js/test_class_fields_advanced.rs

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
    process(x) { return this.#transform(x); }
}
const p = new Processor();
__check(__line(p.process(21)), "42");
let threw = false;
try { p.#transform(1); } catch { threw = true; }
__check(__line(threw), "true");
