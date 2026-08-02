// vybe-test: js/label_break_continue/continue_in_while_loop
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

let i = 0, sum = 0;
while (i < 10) {
    i++;
    if (i % 2 === 0) continue;
    sum += i;
}
console.log(sum); // 1+3+5+7+9 = 25
