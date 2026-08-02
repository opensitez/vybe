// vybe-test: js/class_constructor_errors/derived_super_with_args_passes_to_base
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

class B{constructor(v){this.base=v;}} class D extends B{constructor(v){super(v*2);}} const d=new D(3);__check(__line(d.base), "6");
