// vybe-test: js/bigint_bitwise_operations/bigint_in_switch_strict_match
// origin: languages/js/tests/js/test_bigint_bitwise_operations.rs

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

const v=2n; let r=""; switch(v){case 2n:r="ok";break;default:r="no";} __check(__line(r), "ok");
