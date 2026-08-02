// vybe-test: js/function_prototype_deep/bind_without_partial_args_only_fixes_this
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

function read() { return this.v; } const bound = read.bind({ v: "ok" }); __check(__line(bound()), "ok");
