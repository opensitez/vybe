// vybe-test: js/error_hierarchy/error_has_name_message_stack
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

const e = new TypeError("bad type");
__check(__line(e.name), "TypeError");
__check(__line(e.message), "bad type");
__check(__line(typeof e.stack), "string");
