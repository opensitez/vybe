// vybe-test: js/iterator_protocol/iterator_next_value_is_undefined_after_done
// origin: languages/js/tests/js/test_iterator_protocol.rs

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

function* gen() { yield 1; }
const g = gen();
g.next();
const r = g.next();
__check(__line(r.done), "true");
__check(__line(r.value), "undefined");
