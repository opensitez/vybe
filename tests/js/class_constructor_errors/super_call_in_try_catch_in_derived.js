// vybe-test: js/class_constructor_errors/super_call_in_try_catch_in_derived
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

class B{constructor(){throw new Error("b");}} class D extends B{constructor(){try{super();}catch(e){__check(__line(e.message), "b");}}} try{new D();}catch(e){__check(__line(e instanceof ReferenceError), "true");}
