// vybe-test: js/variadic_bug/variadic_arrow_function_survives_method_call
// origin: languages/js/tests/js/test_variadic_bug.rs

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

const fn = (prefix, ...parts) => {
            const u = prefix.toUpperCase();
            return u + ":" + parts.join(",");
        };
        __check(__line(fn("head", "a", "b")), "HEAD:a,b");
