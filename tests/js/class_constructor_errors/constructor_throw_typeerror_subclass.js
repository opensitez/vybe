// vybe-test: js/class_constructor_errors/constructor_throw_typeerror_subclass
// origin: languages/js/tests/js/test_class_constructor_errors.rs

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

try{new (class{constructor(){throw new TypeError("bad");}})();}catch(e){__check(__line(e.name), "TypeError");}
