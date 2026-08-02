// vybe-test: js/class_constructor_errors/base_constructor_returns_object_derived_gets_it
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

const repl={x:1}; class B{constructor(){return repl;}} class D extends B{constructor(){const o=super();__check(__line(o===repl), "true");}} new D();
