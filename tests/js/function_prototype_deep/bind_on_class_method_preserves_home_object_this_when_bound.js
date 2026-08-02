// vybe-test: js/function_prototype_deep/bind_on_class_method_preserves_home_object_this_when_bound
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

class Box { constructor() { this.v = 1; } read() { return this.v; } } const b = new Box(); const read = b.read.bind(b); __check(__line(read()), "1");
