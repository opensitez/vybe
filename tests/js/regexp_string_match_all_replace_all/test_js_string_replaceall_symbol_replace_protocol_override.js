// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_replaceall_symbol_replace_protocol_override
// origin: languages/js/tests/js/test_js_regexp_string_match_all_replace_all.rs

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

const customReplacer = {
    [Symbol.replace](string, replacement) {
        return `Overridden:${replacement}`;
    }
};
__check(__line("input".replaceAll(customReplacer, "ReplacementText")), "Overridden:ReplacementText");
