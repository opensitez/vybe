// vybe-test: js/iterator_protocol_deep/iterator_result_done_false_value
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

function* gen() { return "final"; }
const it = gen();
const r1 = it.next();
__check(__line(r1.done), "true");
__check(__line(r1.value), "final");
const r2 = it.next();
__check(__line(r2.done), "true");
__check(__line(r2.value), "undefined");
