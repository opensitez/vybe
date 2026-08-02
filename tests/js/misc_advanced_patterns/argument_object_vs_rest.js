// vybe-test: js/misc_advanced_patterns/argument_object_vs_rest
// origin: languages/js/tests/js/test_misc_advanced_patterns.rs

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

function withArgs() {
    return [...arguments].map(x => x * 2);
}
function withRest(...args) {
    return args.map(x => x * 2);
}
__check(__line(withArgs(1, 2, 3).join(",")), "2,4,6");
__check(__line(withRest(1, 2, 3).join(",")), "2,4,6");
__check(__line(Array.isArray(withRest())), "true");
