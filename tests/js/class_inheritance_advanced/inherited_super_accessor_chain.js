// vybe-test: js/class_inheritance_advanced/inherited_super_accessor_chain
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
    set value(raw) {
        this._value = `base:${raw}`;
    }
    get value() {
        return this._value;
    }
}

class Child extends Base {
    set value(raw) {
        super.value = `child:${raw}`;
    }
    get value() {
        return `child->${super.value}`;
    }
}

const c = new Child();
c.value = "payload";
__check(__line(c.value), "child->base:child:payload");
__check(__line(c._value), "base:child:payload");
