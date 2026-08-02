// vybe-test: js/regexp_string_match_all_replace_all/test_js_string_matchall_iterator_protocol_done_state
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

const iter = "x".matchAll(/x/g);
const s1 = iter.next();
const s2 = iter.next();
__check(__line(`${s1.value[0]}|done=${s1.done}`), "x|done=false");
__check(__line(`${s2.value}|done=${s2.done}`), "undefined|done=true");
