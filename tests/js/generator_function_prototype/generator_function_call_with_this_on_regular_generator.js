// vybe-test: js/generator_function_prototype/generator_function_call_with_this_on_regular_generator
// origin: languages/js/tests/js/test_generator_function_prototype.rs

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

function* read() { yield this.v; } const iter = read.call({ v: 9 }); __check(__line(iter.next().value), "9");
