// vybe-test: js/variadic_bug/variadic_args_survive_through_char_at_only
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

function f(fmt, ...args) {
            const c = fmt.charAt(0);
            return [args.length, args[0], c];
        }
        const r = f("=", "X");
        __check(__line(r[0], r[1], r[2]), "1 X =");
