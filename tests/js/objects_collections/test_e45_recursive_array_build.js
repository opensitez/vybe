// vybe-test: js/objects_collections/test_e45_recursive_array_build
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

function range(n) {
            if (n <= 0) return [];
            let arr = range(n - 1);
            arr.push(n);
            return arr;
        }
        __check(__line(range(5).join(",")), "1,2,3,4,5");
