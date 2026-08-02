// vybe-test: js/objects_collections/test_a08_computed_property
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

let key = "color";
        let obj = { [key]: "red" };
        __check(__line(obj.color), "red");
