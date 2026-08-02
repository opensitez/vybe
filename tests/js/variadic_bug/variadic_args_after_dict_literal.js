// vybe-test: js/variadic_bug/variadic_args_after_dict_literal
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

function f(first, ...rest) {
            const obj = { a: 1, b: 2, c: 3 };
            return [rest.length, rest[0], Object.keys(obj).length];
        }
        const r = f("X", "Y", "Z");
        __check(__line(r[0], r[1], r[2]), "2 Y 3");
