// vybe-test: js/abstract_equality_type_coercion_rules/test_js_abstract_equality_date_object_compares_with_timestamp_number
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

const epoch = new Date(0);
__check(__line(epoch == 0), "true");
__check(__line(epoch == 1), "false");
__check(__line(new Date(1000) == 1000), "true");
