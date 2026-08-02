// vybe-test: js/relational_in_instanceof_less_greater_operators/test_js_in_operator_proxy_has_trap
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

const p = new Proxy({}, {
    has(target, prop) {
        return prop === "secret";
    }
});
__check(__line(`${"secret" in p}:${"other" in p}`), "true:false");
