// vybe-test: js/new_collection_methods/set_issupersetof_checks_reverse_containment
// origin: languages/js/tests/js/test_new_collection_methods.rs

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

__check(__line(new Set([1, 2, 3, 4]).isSupersetOf(new Set([2, 3]))), "true");
__check(__line(new Set([1, 2, 3]).isSupersetOf(new Set([2, 4]))), "false");
