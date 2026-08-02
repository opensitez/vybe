// vybe-test: js/spread_rest_advanced/spread_insert_middle
// origin: languages/js/tests/js/test_spread_rest_advanced.rs

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

const start = [1, 2];
const end = [5, 6];
const mid = [3, 4];
const all = [...start, ...mid, ...end];
__check(__line(all.join(",")), "1,2,3,4,5,6");
