// vybe-test: js/error_hierarchy/error_prototype_name_message_defaults
// origin: languages/js/tests/js/test_error_hierarchy.rs

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

const e = new Error();
__check(__line(e.name), "Error");
__check(__line(e.message), "");
