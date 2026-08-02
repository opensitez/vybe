// vybe-test: js/function_invocation_matrix/bound_arrow_ignores_bound_receiver
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

const make = function () {
    const arrow = () => this.name;
    return arrow.bind({ name: "bound" });
};
__check(__line(make.call({ name: "outer" })()), "outer");
