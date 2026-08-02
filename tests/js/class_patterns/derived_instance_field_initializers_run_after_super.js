// vybe-test: js/class_patterns/derived_instance_field_initializers_run_after_super
// origin: languages/js/tests/js/test_class_patterns.rs

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

class Base {
    constructor(seed) { this.seed = seed; }
}

class Derived extends Base {
    doubled = this.seed * 2;
    constructor(seed) {
        super(seed);
        this.seed = seed;
    }
}

const d = new Derived(7);
__check(__line(d.seed), "7");
__check(__line(d.doubled), "14");
