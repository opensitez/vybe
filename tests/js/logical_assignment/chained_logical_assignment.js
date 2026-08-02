// vybe-test: js/logical_assignment/chained_logical_assignment
// origin: languages/js/tests/js/test_logical_assignment.rs

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

let a = null, b = null, c = "found";
a ??= b ??= c;
__check(__line(a), "found");
__check(__line(b), "found");
