// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_no_setter_call_when_short_circuited
// origin: languages/js/tests/js/test_js_logical_assignment_and_or_nullish_operators.rs

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

let setterCount = 0;
const obj = {
    _val: "Initial",
    get val() { return this._val; },
    set val(v) { setterCount++; this._val = v; }
};
obj.val ||= "NewValue"; // Initial is truthy -> short circuits, setter NOT called!
__check(__line(obj.val + "|Setters=" + setterCount), "Initial|Setters=0");
