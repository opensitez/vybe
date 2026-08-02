// vybe-test: js/objects_collections/test_e43_function_pushes_to_array
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

function addItem(arr, item) { arr.push(item); }
        let list = [1, 2];
        addItem(list, 3);
        addItem(list, 4);
        __check(__line(list.join(",")), "1,2,3,4");
