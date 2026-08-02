// vybe-test: js/function_prototype_deep/call_with_explicit_this_mutates_receiver
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

const bag = { count: 0 }; function inc() { this.count++; } inc.call(bag); inc.call(bag); __check(__line(bag.count), "2");
