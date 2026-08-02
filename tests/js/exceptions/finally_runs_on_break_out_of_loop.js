// vybe-test: js/exceptions/finally_runs_on_break_out_of_loop
// origin: languages/js/tests/js/test_exceptions.rs

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

const log = [];
for (let i = 0; i < 3; i++) {
    try {
        if (i === 1) break;
        log.push("try" + i);
    } finally {
        log.push("fin" + i);
    }
}
console.log(log.join(","));
