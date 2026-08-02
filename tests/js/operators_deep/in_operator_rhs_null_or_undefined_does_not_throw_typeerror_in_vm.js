// vybe-test: js/operators_deep/in_operator_rhs_null_or_undefined_does_not_throw_typeerror_in_vm
// origin: languages/js/tests/js/test_operators_deep.rs

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

let nullCheck = false;
let undefinedCheck = false;
try {
    "x" in null;
} catch (e) {
    nullCheck = e instanceof TypeError;
}
try {
    "x" in undefined;
} catch (e) {
    undefinedCheck = e instanceof TypeError;
}
__check(__line(`${nullCheck}:${undefinedCheck}`), "false:false");
