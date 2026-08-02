// vybe-test: js/ecma_operators/prefix_and_postfix_increment_on_accessor_property
// origin: languages/js/tests/js/test_ecma_operators.rs

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
    _val: 10,
    get val() { return this._val; },
    set val(v) { this._val = v; }
};
__check(__line(++obj.val), "11");
__check(__line(obj.val), "11");
__check(__line(obj.val++), "11");
__check(__line(obj.val), "12");
