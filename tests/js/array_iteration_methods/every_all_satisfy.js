// vybe-test: js/array_iteration_methods/every_all_satisfy
// origin: languages/js/tests/js/test_array_iteration_methods.rs

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

__check(__line([2, 4, 6, 8].every(n => n % 2 === 0)), "true");
__check(__line([2, 3, 6].every(n => n % 2 === 0)), "false");
