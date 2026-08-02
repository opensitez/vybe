// vybe-test: js/class_constructor_errors/base_constructor_throw_prevents_derived_field_init
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

try{class B{constructor(){throw new Error("base");}} class D extends B{constructor(){super();this.d=1;}} new D();}catch(e){__check(__line(e.message), "base");}
