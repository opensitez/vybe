// vybe-test: js/object_from_entries/from_entries_with_generator
// origin: languages/js/tests/js/test_object_from_entries.rs

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

function* makeEntries() {
    yield ["x", 1];
    yield ["y", 2];
    yield ["z", 3];
}
const obj = Object.fromEntries([...makeEntries()]);
__check(__line(obj.x), "1");
__check(__line(obj.z), "3");
