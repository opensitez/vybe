// vybe-test: js/iterator_protocol_deep/well_formed_iterator_always_returns_object
// origin: languages/js/tests/js/test_iterator_protocol_deep.rs

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

// next() must always return {value, done}
function* g() { yield 1; }
const it = g();
const r = it.next();
__check(__line(typeof r), "object");
__check(__line("value" in r), "true");
__check(__line("done" in r), "true");
