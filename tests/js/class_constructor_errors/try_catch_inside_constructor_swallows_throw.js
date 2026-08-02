// vybe-test: js/class_constructor_errors/try_catch_inside_constructor_swallows_throw
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

class C{constructor(){try{throw new Error("in");}catch(e){this.msg=e.message;}}} const c=new C();__check(__line(c.msg), "in");
