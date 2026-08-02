// vybe-test: js/label_break_continue/labeled_break_exits_labeled_block
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

let x = 0;
block: {
    x = 1;
    break block;
    x = 2; // never reached
}
__check(__line(x), "1");
