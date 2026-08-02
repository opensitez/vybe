// vybe-test: js/class_constructor_errors/constructor_return_existing_instance_of_same_class
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

class C{constructor(){if(C.cache)return C.cache;C.cache=this;}} C.cache=null; const a=new C(); const b=new C();__check(__line(a===b), "true");
