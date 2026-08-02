// vybe-test: js/object_assign_edge/assign_to_freeze_throws
// origin: languages/js/tests/js/test_object_assign_edge.rs

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

const frozen = Object.freeze({ a: 1 });
frozen.a = 99;  // silently ignored on frozen
__check(__line(frozen.a), "1");
