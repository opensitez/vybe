// vybe-test: js/scope_prototype/prototype_setter_uses_receiver_instance
// origin: languages/js/tests/js/test_scope_prototype.rs

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

const base = {
    set value(v) {
        this._baseValue = v;
    },
    get value() {
        return this._baseValue;
    }
};
const obj = Object.create(base);

obj.value = 9;
__check(__line(obj.value), "9");
__check(__line(base.value), "undefined");
__check(__line(base._baseValue), "undefined");
__check(__line(obj.hasOwnProperty("value")), "false");
