// vybe-test: js/regexp_multiline_m_and_dotall_s_flags/test_js_regexp_global_g_flag_stateful_exec
// origin: languages/js/tests/js/test_js_regexp_multiline_m_and_dotall_s_flags.rs

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

const re = /a/g;
const str = "aba";
const m1 = re.exec(str);
const m2 = re.exec(str);
const m3 = re.exec(str);
__check(__line(`${m1.index}:${m2.index}:${m3 === null}`), "0:2:true");
