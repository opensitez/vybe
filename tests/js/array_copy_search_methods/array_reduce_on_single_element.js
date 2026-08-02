// vybe-test: js/array_copy_search_methods/array_reduce_on_single_element
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

__check(__line([5].reduce((a,b)=>a+b)), "5");
