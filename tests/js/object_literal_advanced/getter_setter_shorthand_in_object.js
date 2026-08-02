// vybe-test: js/object_literal_advanced/getter_setter_shorthand_in_object
// origin: languages/js/tests/js/test_object_literal_advanced.rs

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

const obj = {
    _val: 0,
    get val() { return this._val; },
    set val(v) { this._val = v > 0 ? v : 0; }
};
obj.val = 5;
__check(__line(obj.val), "5");
obj.val = -1;
__check(__line(obj.val), "0");
