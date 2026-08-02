// vybe-test: js/logical_assignment_and_or_nullish_operators/test_js_logical_assignment_accessor_setter_getter_calls
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

let getCount = 0;
let setCount = 0;
const obj = {
    _value: null,
    get value() {
        getCount++;
        return this._value;
    },
    set value(v) {
        setCount++;
        this._value = v;
    }
};

obj.value ||= "fallback";
obj.value ||= "ignored";
obj.value &&= "updated";

__check(__line(`${obj.value}|${getCount}|${setCount}`), "updated|4|3");
