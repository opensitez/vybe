// vybe-test: js/spread_rest_deep/spread_clone_array
// origin: languages/js/tests/js/test_spread_rest_deep.rs

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

const original = [1, 2, 3];
const clone = [...original];
clone.push(4);
__check(__line(original.length), "3");
__check(__line(clone.length), "4");
