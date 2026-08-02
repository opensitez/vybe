// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_in_operator_array_indices
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

const arr = [10, , 30];
__check(__line(`${0 in arr}:${1 in arr}:${2 in arr}:${3 in arr}`), "true:false:true:false");
