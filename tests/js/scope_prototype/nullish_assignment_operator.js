// vybe-test: js/scope_prototype/nullish_assignment_operator
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let a = null;
a ??= 42;
__check(__line(a), "42");
a ??= 99;
__check(__line(a), "42");
