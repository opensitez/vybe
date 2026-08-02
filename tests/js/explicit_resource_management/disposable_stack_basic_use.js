// vybe-test: js/explicit_resource_management/disposable_stack_basic_use
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

const log = [];
{
    using stack = new DisposableStack();
    stack.defer(() => log.push("cleanup1"));
    stack.defer(() => log.push("cleanup2"));
    log.push("work");
}
__check(__line(log.join(",")), "work,cleanup2,cleanup1");
