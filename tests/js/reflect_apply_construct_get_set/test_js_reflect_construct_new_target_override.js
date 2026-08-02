// vybe-test: js/reflect_apply_construct_get_set/test_js_reflect_construct_new_target_override
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set.rs

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
    constructor() {
        this.targetName = new.target.name;
    }
}
class CustomTarget {}

const obj = Reflect.construct(Base, [], CustomTarget);
__check(__line(obj.targetName + "|isCustom=" + (obj instanceof CustomTarget)), "CustomTarget|isCustom=true");
