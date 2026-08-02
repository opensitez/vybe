// vybe-test: js/error_handling_advanced/try_catch_in_promise_chain
// origin: languages/js/tests/js/test_error_handling_advanced.rs

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
    const result = await Promise.resolve(1)
        .then(v => { throw new Error("fail"); })
        .catch(e => "caught: " + e.message)
        .then(v => v + " recovered");
    console.log(result);
}
main();
