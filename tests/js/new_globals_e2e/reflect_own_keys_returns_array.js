// vybe-test: js/new_globals_e2e/reflect_own_keys_returns_array
// origin: languages/js/tests/js/test_new_globals_e2e.rs

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

const o = { a: 1, b: 2, c: 3 };
        const keys = Reflect.ownKeys(o);
        __check(__line(keys.length), "3");
