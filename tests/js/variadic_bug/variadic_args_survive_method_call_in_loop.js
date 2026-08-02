// vybe-test: js/variadic_bug/variadic_args_survive_method_call_in_loop
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
            let out = "";
            let i = 0;
            const len = fmt.length;
            while (i < len) {
                const c = fmt.charAt(i);
                out += c;
                i++;
            }
            return [args.length, args[0]];
        }
        const r = f("=%s=", "X");
        console.log(r[0], r[1]);
