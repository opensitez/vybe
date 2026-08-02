// vybe-test: js/function_invocation_matrix/borrowed_object_method_operates_on_foreign_object
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

const source = {
    x: 10,
    get() {
        return this.x;
    }
};
const target = { x: 22 };
__check(__line(source.get.call(target)), "22");
