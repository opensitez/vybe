// vybe-test: js/ecma_objects/object_entries_preserve_insertion_order
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

const obj = {};
obj.first = 1;
obj.second = 2;
obj.third = 3;
__check(__line(Object.entries(obj).map(([k, v]) => k + ":" + v).join(",")), "first:1,second:2,third:3");
