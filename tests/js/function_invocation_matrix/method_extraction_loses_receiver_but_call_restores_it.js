// vybe-test: js/function_invocation_matrix/method_extraction_loses_receiver_but_call_restores_it
// origin: languages/js/tests/js/test_function_invocation_matrix.rs

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

const obj = {
    name: "Ada",
    speak() {
        return this.name;
    }
};
const loose = obj.speak;
__check(__line(loose.call(obj)), "Ada");
