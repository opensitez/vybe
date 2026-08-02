// vybe-test: js/array_copy_search_methods/array_filter_creates_new_array
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

const a=[1,2,3]; const f=a.filter(x=>x>1); __check(__line(f.join(",")), "2,3");__check(__line(a.length), "3");
