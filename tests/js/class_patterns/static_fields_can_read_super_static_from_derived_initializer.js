// vybe-test: js/class_patterns/static_fields_can_read_super_static_from_derived_initializer
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
    static count = 1;
    static {
        this.count += 1;
    }
}

class Child extends Base {
    static childCount = super.count + 1;
    static {
        this.childCount += Base.count;
    }
}

__check(__line(Base.count), "2");
__check(__line(Child.childCount), "5");
