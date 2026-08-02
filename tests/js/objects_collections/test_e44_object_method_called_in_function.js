// vybe-test: js/objects_collections/test_e44_object_method_called_in_function
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

let obj = { greet: function() { return "hello"; } };
        function callGreet(o) { return o.greet(); }
        __check(__line(callGreet(obj)), "hello");
