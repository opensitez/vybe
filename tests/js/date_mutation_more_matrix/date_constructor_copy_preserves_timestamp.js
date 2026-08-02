// vybe-test: js/date_mutation_more_matrix/date_constructor_copy_preserves_timestamp
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

const a = new Date(1234);
const b = new Date(a);
__check(__line(a.getTime() === b.getTime()), "true");
