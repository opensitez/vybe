// vybe-test: js/for_in_deep/for_in_string_indices
// origin: languages/js/tests/js/test_for_in_deep.rs

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

const keys = [];
const obj = new String("abc");
for (const k in obj) {
    if (/^\d+$/.test(k)) keys.push(k);
}
console.log(keys.join(","));
