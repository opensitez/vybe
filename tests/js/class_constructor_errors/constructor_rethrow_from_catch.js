// vybe-test: js/class_constructor_errors/constructor_rethrow_from_catch
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

try{class C{constructor(){try{throw new Error("a");}catch(e){throw e;}}} new C();}catch(e){__check(__line(e.message), "a");}
