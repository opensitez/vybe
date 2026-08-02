// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_in_operator_own_and_inherited_properties
// origin: languages/js/tests/js/test_js_relational_in_instanceof_less_greater_operators.rs

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

const proto = { parentKey: 1 };
const obj = Object.create(proto);
obj.ownKey = 2;

__check(__line(`${"ownKey" in obj}:${"parentKey" in obj}:${"missing" in obj}`), "true:true:false");
