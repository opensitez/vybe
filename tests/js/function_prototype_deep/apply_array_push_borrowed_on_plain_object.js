// vybe-test: js/function_prototype_deep/apply_array_push_borrowed_on_plain_object
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

const obj = { length: 0 }; Array.prototype.push.apply(obj, ["a", "b"]); __check(__line(obj.length), "2");
