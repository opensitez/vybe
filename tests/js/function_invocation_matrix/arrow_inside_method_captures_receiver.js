// vybe-test: js/function_invocation_matrix/arrow_inside_method_captures_receiver
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
    value: 2,
    run() {
        const f = () => this.value + 1;
        __check(__line(f()), "3");
    }
};
obj.run();
