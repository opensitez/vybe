// vybe-test: js/label_break_continue/labeled_continue_from_inner_try_catch
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

let log = [];
outer: for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) {
            throw "skip";
        }
        log.push("ok" + i);
    } catch {
        log.push("catch" + i);
        continue outer;
    }
    log.push("done" + i);
}
console.log(log.join("|"));
