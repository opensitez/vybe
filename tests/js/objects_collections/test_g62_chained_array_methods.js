// vybe-test: js/objects_collections/test_g62_chained_array_methods
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

let result = [1, 2, 3, 4, 5, 6]
            .filter(x => x % 2 === 0)
            .map(x => x * 10)
            .join(",");
        __check(__line(result), "20,40,60");
