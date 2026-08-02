// vybe-test: js/object_methods_deep/object_values_returns_values
// origin: languages/js/tests/js/test_object_methods_deep.rs

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

const obj = { x: 10, y: 20, z: 30 };
const vals = Object.values(obj);
__check(__line(vals.sort((a,b) => a-b).join(",")), "10,20,30");
