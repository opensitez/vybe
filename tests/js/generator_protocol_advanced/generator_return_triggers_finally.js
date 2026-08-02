// vybe-test: js/generator_protocol_advanced/generator_return_triggers_finally
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

function* gen() {
    try {
        yield 1;
        yield 2;
    } finally {
        yield "cleanup";
    }
}
const g = gen();
g.next(); // advance to yield 1
const r = g.return("early");
// return causes finally to run, yielding "cleanup"
__check(__line(r.value), "cleanup");
__check(__line(r.done), "false");
