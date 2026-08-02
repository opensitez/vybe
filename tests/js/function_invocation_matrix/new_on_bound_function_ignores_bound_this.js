// vybe-test: js/function_invocation_matrix/new_on_bound_function_ignores_bound_this
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

function Person(name) {
    this.name = name;
}
const Bound = Person.bind({ name: "ignored" });
const p = new Bound("Ada");
__check(__line(p.name), "Ada");
