// vybe-test: js/generator_advanced_patterns/generator_return_executes_finally_block
// origin: languages/js/tests/js/test_generator_advanced_patterns.rs

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

function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        __check(__line("finally"), "finally");
    }
}
const g = gen();
g.next();
const r = g.return("done");
__check(__line(r.value + "|" + r.done), "done|true");
