// vybe-test: js/operator_misc/ternary_chained
// origin: languages/js/tests/js/test_operator_misc.rs

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

function grade(n) {
    return n >= 90 ? "A" : n >= 80 ? "B" : n >= 70 ? "C" : "F";
}
__check(__line(grade(95)), "A");
__check(__line(grade(82)), "B");
__check(__line(grade(50)), "F");
