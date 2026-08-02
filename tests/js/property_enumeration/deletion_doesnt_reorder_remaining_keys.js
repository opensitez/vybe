// vybe-test: js/property_enumeration/deletion_doesnt_reorder_remaining_keys
// origin: languages/js/tests/js/test_property_enumeration.rs

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

const obj = { a: 1, b: 2, c: 3, d: 4 };
delete obj.b;
__check(__line(Object.keys(obj).join(",")), "a,c,d");
