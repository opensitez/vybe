// vybe-test: js/ecma_arrays/every_short_circuits_after_failure
// origin: languages/js/tests/js/test_ecma_arrays.rs

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

let seen = [];
const result = [2, 4, 5, 6].every(x => {
    seen.push(x);
    return x % 2 === 0;
});
__check(__line(result), "false");
__check(__line(seen.join(",")), "2,4,5");
