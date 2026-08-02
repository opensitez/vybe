// vybe-test: js/destructuring_advanced/destructure_in_for_of
// origin: languages/js/tests/js/test_destructuring_advanced.rs

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

const pairs = [["a", 1], ["b", 2]];
const keys = [];
for (const [k] of pairs) keys.push(k);
console.log(keys.join(","));
