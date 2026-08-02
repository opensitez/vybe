// vybe-test: js/objects_collections/test_d32_array_some_every
// origin: languages/js/tests/js/js_objects_collections_test.rs

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
        __check(__line([2, 4, 6].every(x => x % 2 === 0)), "true");
        __check(__line([2, 3, 6].every(x => x % 2 === 0)), "false");
