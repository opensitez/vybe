// vybe-test: js/class_constructor_errors/abstract_class_cannot_instantiate_if_runtime_checks
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

class Abstract{constructor(){if(new.target===Abstract)throw new Error("abstract");}} try{new Abstract();}catch(e){__check(__line(e.message), "abstract");}
