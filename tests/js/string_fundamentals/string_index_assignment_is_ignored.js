// vybe-test: js/string_fundamentals/string_index_assignment_is_ignored
// origin: languages/js/tests/js/test_string_fundamentals.rs

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

const name = "abc";
__check(__line(name[0]), "a");
name[0] = "z";
__check(__line(name), "abc");
__check(__line(name.at(0)), "a");
