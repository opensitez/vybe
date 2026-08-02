// vybe-test: js/string_es2022/string_trimstart_trimend_aliases
// origin: languages/js/tests/js/test_string_es2022.rs

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

const s = "  test  ";
__check(__line(s.trimStart() === s.trimLeft()), "true");
__check(__line(s.trimEnd() === s.trimRight()), "true");
