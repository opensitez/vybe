// vybe-test: js/scope_tdz/const_tdz_in_block
// origin: languages/js/tests/js/test_scope_tdz.rs

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

let result;
{
    try {
        result = x;
        const x = 1;
    } catch (e) {
        result = "tdz";
    }
}
__check(__line(result), "tdz");
