// vybe-test: js/date_arithmetic_edge/date_tojson_is_quoted_iso
// origin: languages/js/tests/js/test_date_arithmetic_edge.rs

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

__check(__line(JSON.parse(JSON.stringify(new Date(0))).startsWith("1970")), "true");
