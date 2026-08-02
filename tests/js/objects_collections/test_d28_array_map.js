// vybe-test: js/objects_collections/test_d28_array_map
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

let doubled = [1, 2, 3].map(x => x * 2);
        __check(__line(doubled.join(",")), "2,4,6");
