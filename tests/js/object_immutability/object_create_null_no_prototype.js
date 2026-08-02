// vybe-test: js/object_immutability/object_create_null_no_prototype
// origin: languages/js/tests/js/test_object_immutability.rs

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

const safe = Object.create(null);
safe.key = "value";
// No toString, no hasOwnProperty — truly property-free
__check(__line(typeof safe.toString), "undefined");
__check(__line(safe.key), "value");
