// vybe-test: js/class_fields_advanced/static_private_field
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

class IdGenerator {
    static #nextId = 0;
    static generate() { return ++IdGenerator.#nextId; }
}
__check(__line(IdGenerator.generate()), "1");
__check(__line(IdGenerator.generate()), "2");
__check(__line(IdGenerator.generate()), "3");
