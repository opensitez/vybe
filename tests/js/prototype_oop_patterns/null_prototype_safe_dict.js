// vybe-test: js/prototype_oop_patterns/null_prototype_safe_dict
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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

const dict = Object.create(null);
dict.constructor = "fake";
dict.toString = "fake";
dict.hasOwnProperty = "fake";
// Null prototype means no inherited methods
__check(__line(Object.getPrototypeOf(dict)), "null");
__check(__line("constructor" in dict), "true");
// Object.hasOwn works on null-prototype
dict.real = 42;
__check(__line(Object.hasOwn(dict, "real")), "true");
