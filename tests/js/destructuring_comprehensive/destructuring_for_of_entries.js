// vybe-test: js/destructuring_comprehensive/destructuring_for_of_entries
// origin: languages/js/tests/js/test_destructuring_comprehensive.rs

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

const map = new Map([["a", 1], ["b", 2], ["c", 3]]);
const results = [];
for (const [key, value] of map) results.push(`${key}=${value}`);
console.log(results.join(","));
