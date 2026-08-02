// vybe-test: js/ecma_objects/object_from_entries_overwrites_duplicate_keys
// origin: languages/js/tests/js/test_ecma_objects.rs

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

const obj = Object.fromEntries([["a", 1], ["a", 2], ["b", 3]]);
__check(__line(obj.a), "2");
__check(__line(obj.b), "3");
