// vybe-test: js/label_break_continue/triple_nested_labeled
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

let count = 0;
a: for (let i = 0; i < 3; i++) {
    b: for (let j = 0; j < 3; j++) {
        c: for (let k = 0; k < 3; k++) {
            if (k === 1) continue b;
            count++;
        }
    }
}
console.log(count); // each i,j pair contributes 1 (k=0), skips rest
