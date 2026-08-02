// vybe-test: js/objects_collections/test_e48_closure_captures_object
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

let obj = { count: 0 };
        let increment = () => { obj.count = obj.count + 1; };
        increment();
        increment();
        increment();
        __check(__line(obj.count), "3");
