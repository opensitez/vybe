// vybe-test: js/exceptions/throw_bigint_is_caught_by_value
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

let caught;
try { throw 10n; } catch (e) { caught = e; }
__check(__line(caught === 10n), "true");
__check(__line(typeof caught), "bigint");
