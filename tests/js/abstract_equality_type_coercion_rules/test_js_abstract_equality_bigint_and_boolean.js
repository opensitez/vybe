// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_bigint_and_boolean
// origin: languages/js/tests/js/test_js_abstract_equality_type_coercion_rules.rs

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

__check(__line(`${0n == false}:${1n == true}:${2n == true}:${0n === false}`), "true:true:false:false");
