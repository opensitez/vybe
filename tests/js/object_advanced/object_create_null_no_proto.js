// vybe-test: js/object_advanced/object_create_null_no_proto
// origin: languages/js/tests/js/test_object_advanced.rs

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

let obj = Object.create(null);
obj.x = 42;
__check(__line(obj.x), "42");
__check(__line(obj.toString), "undefined");
