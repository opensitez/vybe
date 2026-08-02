// vybe-test: js/arrow_functions/arrow_lexical_super_in_method
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

class B{who(){return "base";}} class D extends B{who(){const a=()=>super.who();return a()+"+d";}} __check(__line(new D().who()), "base+d");
