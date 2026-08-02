// vybe-test: js/class_constructor_errors/class_instance_custom_tostringtag_brand
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

class Custom{[Symbol.toStringTag]="MyCustomClass";} __check(__line(Object.prototype.toString.call(new Custom())), "[object MyCustomClass]");
