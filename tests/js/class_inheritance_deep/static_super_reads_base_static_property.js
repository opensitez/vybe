// vybe-test: js/class_inheritance_deep/static_super_reads_base_static_property
// origin: languages/js/tests/js/test_class_inheritance_deep.rs

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
    static marker = "base";
}
class Child extends Base {
    static marker = "child";
    static getBaseMarker() {
        return super.marker;
    }
}
__check(__line(Child.marker), "child");
__check(__line(Child.getBaseMarker()), "base");
__check(__line(Base.marker), "base");
