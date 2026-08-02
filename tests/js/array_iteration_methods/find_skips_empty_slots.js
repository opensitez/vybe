// vybe-test: js/array_iteration_methods/find_skips_empty_slots
// origin: languages/js/tests/js/test_array_iteration_methods.rs

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

const arr = [, 1, , 3];
let seen = [];
const value = arr.find((value, index) => {
    seen.push(index);
    return value === 3;
});
__check(__line(value), "3");
__check(__line(seen.join(",")), "0,1,2,3");
