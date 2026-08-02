// vybe-test: js/eval_dynamic_code/typeof_undeclared_safe_check
// origin: languages/js/tests/js/test_eval_dynamic_code.rs

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

// typeof doesn't throw for undeclared
console.log(typeof COMPLETELY_UNDECLARED_VARIABLE);
