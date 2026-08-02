// vybe-test: js/function_invocation_matrix/bind_chained_receiver_stays_from_first_bind
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

function show() {
    return this.name;
}
const one = show.bind({ name: "a" });
const two = one.bind({ name: "b" });
__check(__line(two()), "a");
