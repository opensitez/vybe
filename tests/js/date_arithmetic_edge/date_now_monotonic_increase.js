// vybe-test: js/date_arithmetic_edge/date_now_monotonic_increase
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

const a=Date.now(); const b=Date.now(); __check(__line(b>=a), "true");
