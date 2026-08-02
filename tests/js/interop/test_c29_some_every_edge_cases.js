// vybe-test: js/interop/test_c29_some_every_edge_cases
// origin: languages/js/tests/js/js_interop_test.rs

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

__check(__line([1, 2, 3].some(x => x > 2)), "true");
        __check(__line([1, 2, 3].some(x => x > 10)), "false");
        __check(__line([2, 4, 6].every(x => x % 2 === 0)), "true");
        __check(__line([2, 3, 6].every(x => x % 2 === 0)), "false");
        __check(__line([].some(x => true)), "false");
        __check(__line([].every(x => false)), "true");
