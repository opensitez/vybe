// vybe-test: js/function_prototype_deep/bind_cannot_override_fixed_this_via_call
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

function id() { return this.tag; } const fixed = id.bind({ tag: "one" }); __check(__line(fixed.call({ tag: "two" })), "one");
