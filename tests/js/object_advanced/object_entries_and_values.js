// vybe-test: js/object_advanced/object_entries_and_values
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

const obj = { a: 1, b: 2 };
__check(__line(Object.entries(obj).length), "2");
__check(__line(Object.keys(obj).length), "2");
__check(__line(Object.values(obj).join(",")), "1,2");
