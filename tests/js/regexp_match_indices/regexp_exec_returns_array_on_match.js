// vybe-test: js/regexp_match_indices/regexp_exec_returns_array_on_match
// origin: languages/js/tests/js/test_regexp_match_indices.rs

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

const r=/a(b)c/; const m=r.exec("abc"); __check(__line(m[0]), "abc");__check(__line(m[1]), "b");
