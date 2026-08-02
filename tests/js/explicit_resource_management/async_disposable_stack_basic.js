// vybe-test: js/explicit_resource_management/async_disposable_stack_basic
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
async function main() {
    await using stack = new AsyncDisposableStack();
    stack.defer(async () => {
        await Promise.resolve();
        log.push("cleanup");
    });
    log.push("work");
}
main().then(() => console.log(log.join(",")));
