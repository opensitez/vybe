// vybe-test: js/explicit_resource_management/disposable_stack_adopt
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
    const handle = stack.adopt({ id: 1 }, (h) => log.push("close:" + h.id));
    log.push("use:" + handle.id);
}
__check(__line(log.join(",")), "use:1,close:1");
