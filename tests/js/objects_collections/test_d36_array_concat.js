// vybe-test: js/objects_collections/test_d36_array_concat
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

let a = [1, 2];
        let b = [3, 4];
        let c = a.concat(b);
        __check(__line(c.join(","), a.length), "1,2,3,4 2");
