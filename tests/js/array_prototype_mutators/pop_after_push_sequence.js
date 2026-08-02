// vybe-test: js/array_prototype_mutators/pop_after_push_sequence
// origin: languages/js/tests/js/test_array_prototype_mutators.rs

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

const a=[]; a.push(1); a.push(2); __check(__line(a.pop()), "2"); __check(__line(a.pop()), "1"); __check(__line(a.pop()===undefined), "true");
