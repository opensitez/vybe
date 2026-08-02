// vybe-test: js/error_cause_aggregate/error_name_can_be_customized_on_instance
// origin: languages/js/tests/js/test_error_cause_aggregate.rs

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

const e = new Error("test");
e.name = "CustomError";
__check(__line(e.name), "CustomError");
__check(__line(e.message), "test");
