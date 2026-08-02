// vybe-test: js/label_break_continue/labeled_for_of
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

let found = null;
outer: for (const arr of [[1,2],[3,4],[5,6]]) {
    for (const x of arr) {
        if (x === 4) { found = x; break outer; }
    }
}
console.log(found);
