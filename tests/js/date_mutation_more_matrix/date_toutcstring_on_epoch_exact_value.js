// vybe-test: js/date_mutation_more_matrix/date_toutcstring_on_epoch_exact_value
// origin: languages/js/tests/js/test_date_mutation_more_matrix.rs

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

__check(__line(new Date(0).toUTCString()), "Thu, 01 Jan 1970 00:00:00 GMT");
