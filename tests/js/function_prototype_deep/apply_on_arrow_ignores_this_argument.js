// vybe-test: js/function_prototype_deep/apply_on_arrow_ignores_this_argument
// origin: languages/js/tests/js/test_function_prototype_deep.rs

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

const a = () => 1; const b = () => 1; __check(__line(a.apply({ x: 1 }, []) === b.apply({ x: 2 }, [])), "true");
