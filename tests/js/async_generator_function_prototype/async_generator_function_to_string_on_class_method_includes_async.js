// vybe-test: js/async_generator_function_prototype/async_generator_function_to_string_on_class_method_includes_async
// origin: languages/js/tests/js/test_async_generator_function_prototype.rs

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

class G { async *run() { yield 1; } } __check(__line(Function.prototype.toString.call(G.prototype.run).includes("async")), "true");
