// vybe-test: js/property_ordering/keys_string_properties_insertion_order
// origin: languages/js/tests/js/test_property_ordering.rs

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
obj.c = 3;
obj.a = 1;
obj.b = 2;
__check(__line(Object.keys(obj).join(",")), "c,a,b");
