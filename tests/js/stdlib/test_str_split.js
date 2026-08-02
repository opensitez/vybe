// vybe-test: js/stdlib/test_str_split
// origin: languages/js/tests/js/js_stdlib_test.rs

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

let parts = "a,b,c".split(","); __check(__line(parts[0], parts[1], parts[2]), "a b c")
