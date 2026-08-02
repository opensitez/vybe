// vybe-test: js/object_descriptors/object_assign_multiple_sources
// origin: languages/js/tests/js/test_object_descriptors.rs

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

const target = { a: 1 };
const result = Object.assign(target, { b: 2 }, { c: 3 }, { d: 4 });
__check(__line(result === target), "true");
__check(__line(Object.keys(result).sort().join(",")), "a,b,c,d");
