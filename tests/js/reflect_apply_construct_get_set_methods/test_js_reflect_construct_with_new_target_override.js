// vybe-test: js/reflect_apply_construct_get_set_methods/test_js_reflect_construct_with_new_target_override
// origin: languages/js/tests/js/test_js_reflect_apply_construct_get_set_methods.rs

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

function Base() { this.base = true; }
function Sub() {}
Sub.prototype = Object.create(Base.prototype);
Sub.prototype.subProp = "sub";

const obj = Reflect.construct(Base, [], Sub);
__check(__line(obj.subProp + "|" + (obj instanceof Sub)), "sub|true");
