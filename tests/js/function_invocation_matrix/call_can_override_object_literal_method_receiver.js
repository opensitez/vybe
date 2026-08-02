// vybe-test: js/function_invocation_matrix/call_can_override_object_literal_method_receiver
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

const left = {
    name: "left",
    show() {
        return this.name;
    }
};
const right = { name: "right" };
__check(__line(left.show.call(right)), "right");
