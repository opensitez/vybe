// vybe-test: js/exceptions/throw_symbol_is_caught_by_identity
// origin: languages/js/tests/js/test_exceptions.rs

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

const s = Symbol("boom");
let same = false;
try { throw s; } catch (e) { same = e === s; }
__check(__line(same), "true");
