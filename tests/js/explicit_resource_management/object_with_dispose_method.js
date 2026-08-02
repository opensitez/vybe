// vybe-test: js/explicit_resource_management/object_with_dispose_method
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
const res = {
    [Symbol.dispose]() { log.push("disposed"); }
};
{
    using r = res;
}
__check(__line(log.join(",")), "disposed");
