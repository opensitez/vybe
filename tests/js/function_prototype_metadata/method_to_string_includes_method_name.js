// vybe-test: js/function_prototype_metadata/method_to_string_includes_method_name
// origin: languages/js/tests/js/test_function_prototype_metadata.rs

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

const obj = { ping() { return 1; } }; __check(__line(Function.prototype.toString.call(obj.ping).includes("ping")), "true");
