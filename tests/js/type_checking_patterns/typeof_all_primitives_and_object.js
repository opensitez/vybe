// vybe-test: js/type_checking_patterns/typeof_all_primitives_and_object
// origin: languages/js/tests/js/test_type_checking_patterns.rs

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
__check(__line(typeof "str"), "string");
__check(__line(typeof Symbol()), "symbol");
__check(__line(typeof 1n), "bigint");
__check(__line(typeof function(){}), "function");
__check(__line(typeof {}), "object");
__check(__line(typeof []), "object");
