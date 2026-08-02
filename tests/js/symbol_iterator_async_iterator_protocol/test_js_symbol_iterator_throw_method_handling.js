// vybe-test: js/symbol_iterator_async_iterator_protocol/test_js_symbol_iterator_throw_method_handling
// origin: languages/js/tests/js/test_js_symbol_iterator_async_iterator_protocol.rs

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
    } catch (e) {
        yield "CaughtInGen: " + e.message;
    }
}
const g = gen();
g.next();
__check(__line(g.throw(new Error("ExternalError")).value), "CaughtInGen: ExternalError");
