// vybe-test: js/explicit_resource_management/disposable_stack_move_transfers_ownership
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
let outer;
{
    using stack = new DisposableStack();
    stack.defer(() => log.push("cleanup"));
    outer = stack.move();
    log.push("inner disposed:" + stack.disposed);
}
log.push("outer disposed before:" + outer.disposed);
outer.dispose();
log.push("outer disposed after:" + outer.disposed);
__check(__line(log.join(",")), "inner disposed:true,outer disposed before:false,cleanup,outer disposed after:true");
