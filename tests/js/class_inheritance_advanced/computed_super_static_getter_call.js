// vybe-test: js/class_inheritance_advanced/computed_super_static_getter_call
// origin: languages/js/tests/js/test_class_inheritance_advanced.rs

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
    static get marker() {
        return "base-marker";
    }
}

class Child extends Base {
    static getComputedMarker() {
        const key = "marker";
        return super[key];
    }
}

__check(__line(Child.getComputedMarker()), "base-marker");
