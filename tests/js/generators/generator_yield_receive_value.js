// vybe-test: js/generators/generator_yield_receive_value
// origin: languages/js/tests/js/test_generators.rs

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

function* echo() {
    let msg = yield "ready";
    __check(__line("received: " + msg), "ready");
    yield "done";
}
let g = echo();
__check(__line(g.next().value), "received: hello");
__check(__line(g.next("hello").value), "done");
