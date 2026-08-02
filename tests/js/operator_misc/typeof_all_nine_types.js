// vybe-test: js/operator_misc/typeof_all_nine_types
// origin: languages/js/tests/js/test_operator_misc.rs

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

__check(__line(typeof undefined), "undefined");
__check(__line(typeof null), "object");
__check(__line(typeof true), "boolean");
__check(__line(typeof 42), "number");
__check(__line(typeof "string"), "string");
__check(__line(typeof Symbol()), "symbol");
__check(__line(typeof 42n), "bigint");
__check(__line(typeof function(){}), "function");
__check(__line(typeof {}), "object");
