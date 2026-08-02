// vybe-test: js/exceptions/throw_null_is_caught_as_null
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

let ok = false;
try { throw null; } catch (e) { ok = e === null; }
__check(__line(ok), "true");
