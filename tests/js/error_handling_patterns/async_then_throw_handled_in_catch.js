// vybe-test: js/error_handling_patterns/async_then_throw_handled_in_catch
// origin: languages/js/tests/js/test_error_handling_patterns.rs

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

async function main() {
    try {
        await Promise.resolve().then(() => {
            throw new Error("then_err");
        });
    } catch (e) {
        console.log(e.message);
    }
}
main();
