// vybe-test: js/string_unicode_deep/string_char_at_vs_index
// origin: languages/js/tests/js/test_string_unicode_deep.rs

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

const s = "hello";
console.log(s.charAt(1));
console.log(s[1]);
console.log(s.charAt(99));  // "" for out of bounds
console.log(s[99]);          // undefined for out of bounds
