// vybe-test: js/closure_scope_deep/pipe_left_to_right
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

const pipe = (...fns) => x => fns.reduce((v, f) => f(v), x);
const process = pipe(
    s => s.trim(),
    s => s.toLowerCase(),
    s => s.replace(/\s+/g, "-")
);
__check(__line(process("  Hello World  ")), "hello-world");
