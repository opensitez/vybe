// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_in_operator_numeric_left_operand_coerces_to_property_key
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

const arr = ["a", "b", "c"];
__check(__line(`${1 in arr}:${"1" in arr}`), "true:true");
__check(__line(`${99 in arr}:${99n in arr}`), "false:true");
