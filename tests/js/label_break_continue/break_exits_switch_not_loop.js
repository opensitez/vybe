// vybe-test: js/label_break_continue/break_exits_switch_not_loop
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
for (let i = 0; i < 3; i++) {
    switch (i) {
        case 1: break; // exits switch, not loop
    }
    result.push(i);
}
console.log(result.join(","));
