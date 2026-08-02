// vybe-test: js/getter_setter_deep/computed_getter_name
// origin: languages/js/tests/js/test_getter_setter_deep.rs

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

const prop = "value";
const obj = {
    _v: 42,
    get [prop]() { return this._v; },
    set [prop](v) { this._v = v; }
};
__check(__line(obj.value), "42");
obj.value = 100;
__check(__line(obj.value), "100");
