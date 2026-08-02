// vybe-test: js/generators/generator_rest_args_survive_fresh_throw
// origin: languages/js/tests/js/test_generators.rs

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

function* guarded(head, ...rest) {
    try {
        yield rest.length;
    } catch (err) {
        __check(__line(rest.join(",")), "b,c");
        yield err.message;
    }
}
let g = guarded("a", "b", "c");
let result = g.throw(new Error("stop"));
__check(__line(result.value), "stop");
__check(__line(result.done), "false");
