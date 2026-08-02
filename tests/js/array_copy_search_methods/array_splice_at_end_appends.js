// vybe-test: js/array_copy_search_methods/array_splice_at_end_appends
// origin: languages/js/tests/js/test_array_copy_search_methods.rs

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

const a=[1,2]; a.splice(2,0,3); __check(__line(a.join(",")), "1,2,3");
