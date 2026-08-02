// vybe-test: js/generator_return_throw_next_state_machine/test_js_generator_finally_block_yields_on_return
// origin: languages/js/tests/js/test_js_generator_return_throw_next_state_machine.rs

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
    } finally {
        yield "FinallyYield";
    }
}
const g = gen();
g.next();
const ret = g.return("EarlyRet");
__check(__line(`${ret.value}:${ret.done}`), "FinallyYield:false"); // Yield in finally intercepts return!
