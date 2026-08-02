// vybe-test: js/string_code_point_at_from_code_point_surrogates/test_js_string_for_of_loop_iterates_by_code_points
// origin: languages/js/tests/js/test_js_string_code_point_at_from_code_point_surrogates.rs

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

const str = "A😀B";
const codePoints = [];
for (const char of str) {
    codePoints.push(char.codePointAt(0));
}
console.log(codePoints.join(","));
