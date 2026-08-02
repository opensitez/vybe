// vybe-test: js/interop/test_b13_object_literal_with_methods
// origin: languages/js/tests/js/js_interop_test.rs

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

let obj = {
            name: "test",
            greet() { return "hi from " + this.name; }
        };
        __check(__line(obj.greet()), "hi from test");
