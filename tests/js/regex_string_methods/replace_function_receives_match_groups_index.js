// vybe-test: js/regex_string_methods/replace_function_receives_match_groups_index
// origin: languages/js/tests/js/test_regex_string_methods.rs

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

const result = "hello world".replace(/(\w+)/g, (match, group1, index) => {
    return `[${match}@${index}]`;
});
__check(__line(result), "[hello@0] [world@6]");
