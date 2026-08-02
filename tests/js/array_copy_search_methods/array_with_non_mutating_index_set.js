// vybe-test: js/array_copy_search_methods/array_with_non_mutating_index_set
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

const a=[1,2,3]; const w=a.with(1,9); __check(__line(a[1]), "2");__check(__line(w[1]), "9");
