// vybe-test: js/ecma_iterators/for_of_string_iterates_unicode_code_points
// origin: languages/js/tests/js/test_ecma_iterators.rs

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

let chars = [];
for (const ch of "A😀B") {
    chars.push(ch);
}
console.log(chars.length);
console.log(chars[1]);
