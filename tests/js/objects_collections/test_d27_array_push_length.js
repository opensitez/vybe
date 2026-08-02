// vybe-test: js/objects_collections/test_d27_array_push_length
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

let arr = [1, 2];
        arr.push(3);
        arr.push(4);
        __check(__line(arr.length, arr.join(",")), "4 1,2,3,4");
