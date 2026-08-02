// vybe-test: js/new_collection_methods/set_issubsetof_checks_containment
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

__check(__line(new Set([2, 3]).isSubsetOf(new Set([1, 2, 3, 4]))), "true");
__check(__line(new Set([2, 5]).isSubsetOf(new Set([1, 2, 3, 4]))), "false");
