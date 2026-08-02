// vybe-test: js/arrow_functions/class_field_arrow_binds_instance
// origin: languages/js/tests/js/test_arrow_functions.rs

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

class C{v=5;get=()=>this.v;} const c=new C(); const g=c.get; __check(__line(g()), "5");
