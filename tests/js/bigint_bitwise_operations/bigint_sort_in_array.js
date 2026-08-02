// vybe-test: js/bigint_bitwise_operations/bigint_sort_in_array
// origin: languages/js/tests/js/test_bigint_bitwise_operations.rs

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

__check(__line([3n,1n,2n].sort((a,b)=>a<b?-1:1).join(",")), "1,2,3");
