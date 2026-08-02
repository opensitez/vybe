// vybe-test: js/scope_prototype/symbol_basic
// origin: languages/js/tests/js/test_scope_prototype.rs

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

let s1 = Symbol("desc");
let s2 = Symbol("desc");
__check(__line(typeof s1), "symbol");
__check(__line(s1 === s2), "false");
__check(__line(s1.toString()), "Symbol(desc)");
