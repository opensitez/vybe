// vybe-test: js/type_coercion_deep/symbol_to_number_throws
// origin: languages/js/tests/js/test_type_coercion_deep.rs

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

let threw = false;
try { +Symbol("test"); } catch (e) { threw = e instanceof TypeError; }
__check(__line(threw), "true");
