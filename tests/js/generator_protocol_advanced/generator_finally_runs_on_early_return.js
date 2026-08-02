// vybe-test: js/generator_protocol_advanced/generator_finally_runs_on_early_return
// origin: languages/js/tests/js/test_generator_protocol_advanced.rs

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
function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        log.push("finally");
    }
}
const g = gen();
g.next();
g.return("stop");
__check(__line(log.join(",")), "finally");
