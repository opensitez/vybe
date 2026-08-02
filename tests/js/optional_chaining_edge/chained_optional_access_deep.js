// vybe-test: js/optional_chaining_edge/chained_optional_access_deep
// origin: languages/js/tests/js/test_optional_chaining_edge.rs

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

const a = { b: { c: { d: 42 } } };
__check(__line(a?.b?.c?.d), "42");
__check(__line(a?.b?.x?.d), "undefined");
__check(__line(a?.z?.c?.d), "undefined");
