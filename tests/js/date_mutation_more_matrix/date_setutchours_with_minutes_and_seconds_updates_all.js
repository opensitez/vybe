// vybe-test: js/date_mutation_more_matrix/date_setutchours_with_minutes_and_seconds_updates_all
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

const d = new Date(Date.UTC(2024, 0, 1, 1, 2, 3, 4));
d.setUTCHours(4, 5, 6, 7);
__check(__line(d.getUTCHours()), "4");
__check(__line(d.getUTCMinutes()), "5");
__check(__line(d.getUTCSeconds()), "6");
__check(__line(d.getUTCMilliseconds()), "7");
