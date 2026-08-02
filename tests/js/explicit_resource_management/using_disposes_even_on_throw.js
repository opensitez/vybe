// vybe-test: js/explicit_resource_management/using_disposes_even_on_throw
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

const log = [];
try {
    using r = { [Symbol.dispose]() { log.push("disposed"); } };
    throw new Error("oops");
} catch (e) {
    log.push("caught:" + e.message);
}
__check(__line(log.join(",")), "disposed,caught:oops");
