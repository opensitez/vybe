// vybe-test: js/class_constructor_errors/new_on_bound_class_constructor
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

class C{constructor(v){this.v=v;}} const B=C.bind(null,7); const i=new B();__check(__line(i.v), "7");
