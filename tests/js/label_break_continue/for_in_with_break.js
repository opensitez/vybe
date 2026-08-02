// vybe-test: js/label_break_continue/for_in_with_break
// origin: languages/js/tests/js/test_label_break_continue.rs

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

const obj = { a: 1, b: 2, c: 3 };
let first = null;
for (const key in obj) {
    first = key;
    break;
}
console.log(first);
