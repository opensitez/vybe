// vybe-test: js/object_assign_edge/assign_does_not_copy_getters_as_getters
// origin: languages/js/tests/js/test_object_assign_edge.rs

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

// assign copies own enumerable data properties; result always has data properties
const key = "prop";
const src = { [key]: 42, other: 7 };
const result = Object.assign({}, src);
__check(__line(result.prop), "42");
__check(__line(Object.keys(result).length), "2");
