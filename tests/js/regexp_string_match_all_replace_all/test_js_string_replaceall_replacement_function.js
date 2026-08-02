// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_replaceall_replacement_function
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

const str = "1 2 3 4";
const res = str.replaceAll(/\d/g, match => String(Number(match) * 10));
__check(__line(res), "10 20 30 40");
