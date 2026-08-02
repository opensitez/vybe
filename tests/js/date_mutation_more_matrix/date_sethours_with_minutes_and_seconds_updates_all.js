// vybe-test: js/date_mutation_more_matrix/date_sethours_with_minutes_and_seconds_updates_all
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

const d = new Date(2024, 0, 1, 1, 2, 3);
d.setHours(4, 5, 6, 7);
__check(__line(d.getHours()), "4");
__check(__line(d.getMinutes()), "5");
__check(__line(d.getSeconds()), "6");
__check(__line(d.getMilliseconds()), "7");
