// vybe-test: js/variadic_bug/variadic_args_with_default_parameter_survives_method_call
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

function f(first = "default", ...rest) {
            const u = first.toUpperCase();
            return [u, rest.length, rest[0]];
        }
        const r = f("hello", "X", "Y");
        __check(__line(r[0], r[1], r[2]), "HELLO 2 X");
