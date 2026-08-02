// vybe-test: js/objects_collections/test_g63_reduce_to_object
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

let pairs = [["a", 1], ["b", 2], ["c", 3]];
        let obj = pairs.reduce((acc, pair) => {
            acc[pair[0]] = pair[1];
            return acc;
        }, {});
        __check(__line(obj.a, obj.b, obj.c), "1 2 3");
