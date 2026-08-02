// vybe-test: js/class_constructor_errors/new_target_chain_through_super
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

class A{constructor(){this.chain=new.target.name;}} class B extends A{constructor(){super();}} class C extends B{constructor(){super();}} const c=new C();__check(__line(c.chain), "C");
