// vybe-test: js/object_advanced/object_from_entries
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

const pairs = [["a", 1], ["b", 2], ["c", 3]];
const obj = Object.fromEntries(pairs);
__check(__line(obj.a), "1");
__check(__line(obj.b), "2");
__check(__line(obj.c), "3");
