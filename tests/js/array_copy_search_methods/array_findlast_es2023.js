// vybe-test: js/array_copy_search_methods/array_findlast_es2023
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

__check(__line([1,4,3,4].findLast(x=>x>2)), "4");
