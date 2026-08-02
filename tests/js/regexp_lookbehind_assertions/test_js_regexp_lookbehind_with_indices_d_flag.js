// vybe-test: js/regexp_lookbehind_assertions/test_js_regexp_lookbehind_with_indices_d_flag
// origin: languages/js/tests/js/test_js_regexp_lookbehind_assertions.rs

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

const re = /(?<=\$)\d+/d;
const match = re.exec("Cost: $50");
__check(__line(match.indices[0].join(":")), "7:9"); // Indices match span of digits 50 (index 7..9)
