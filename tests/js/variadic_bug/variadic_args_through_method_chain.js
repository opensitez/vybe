// vybe-test: js/variadic_bug/variadic_args_through_method_chain
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

function f(s, ...args) {
            const r = s.toUpperCase().toLowerCase().trim();
            return [args.length, args[0], r];
        }
        const r = f("  Hi  ", 1, 2);
        __check(__line(r[0], r[1], r[2]), "2 1 hi");
