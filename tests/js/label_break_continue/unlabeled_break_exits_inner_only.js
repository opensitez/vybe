// vybe-test: js/label_break_continue/unlabeled_break_exits_inner_only
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

let result = [];
for (let i = 0; i < 2; i++) {
    for (let j = 0; j < 3; j++) {
        if (j === 1) break;
        result.push(i + "," + j);
    }
}
console.log(result.join("|"));
