// vybe-test: js/bigint_advanced/bigint_to_string_with_radix
// origin: languages/js/tests/js/test_bigint_advanced.rs

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

__check(__line((255n).toString(16)), "ff");
__check(__line((8n).toString(2)), "1000");
