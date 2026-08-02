// vybe-test: js/number_precision_edge/number_parsefloat_basics
// origin: languages/js/tests/js/test_number_precision_edge.rs

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

__check(__line(Number.parseFloat("3.14")), "3.14");
__check(__line(Number.parseFloat("3.14xyz")), "3.14");
__check(__line(isNaN(Number.parseFloat("abc"))), "true");
