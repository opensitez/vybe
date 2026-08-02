// vybe-test: js/misc_es_features/optional_chain_method_call
// origin: languages/js/tests/js/test_misc_es_features.rs

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

const obj = { greet() { return "hello"; } };
__check(__line(obj?.greet()), "hello");
__check(__line(obj?.missing?.()), "undefined");
