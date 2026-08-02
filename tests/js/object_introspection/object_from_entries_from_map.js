// vybe-test: js/object_introspection/object_from_entries_from_map
// origin: languages/js/tests/js/test_object_introspection.rs

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

const map = new Map([["one", 1], ["two", 2], ["three", 3]]);
const obj = Object.fromEntries(map);
__check(__line(obj.one), "1");
__check(__line(obj.two), "2");
__check(__line(obj.three), "3");
