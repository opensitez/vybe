// vybe-test: js/generator_protocol_advanced/generator_return_ends_iteration
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

function* count() {
    yield 1; yield 2; yield 3;
}
const g = count();
g.next(); // 1
const r = g.return("done");
__check(__line(r.value), "done");
__check(__line(r.done), "true");
const next = g.next();
__check(__line(next.done), "true"); // still done
