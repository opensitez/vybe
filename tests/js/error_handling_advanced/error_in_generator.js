// vybe-test: js/error_handling_advanced/error_in_generator
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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
    } catch(e) {
        yield "caught: " + e.message;
    }
    yield 3;
}
const g = gen();
__check(__line(g.next().value), "1");
__check(__line(g.throw(new Error("oops")).value), "caught: oops");
__check(__line(g.next().value), "3");
