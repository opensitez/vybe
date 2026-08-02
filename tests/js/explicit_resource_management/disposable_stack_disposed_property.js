// vybe-test: js/explicit_resource_management/disposable_stack_disposed_property
// origin: languages/js/tests/js/test_explicit_resource_management.rs

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

const stack = new DisposableStack();
__check(__line(stack.disposed), "false");
stack.dispose();
__check(__line(stack.disposed), "true");
