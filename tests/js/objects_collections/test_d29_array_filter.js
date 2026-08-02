// vybe-test: js/objects_collections/test_d29_array_filter
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

let evens = [1, 2, 3, 4, 5, 6].filter(x => x % 2 === 0);
        __check(__line(evens.join(",")), "2,4,6");
