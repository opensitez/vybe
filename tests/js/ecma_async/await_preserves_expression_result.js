// vybe-test: js/ecma_async/await_preserves_expression_result
// origin: languages/js/tests/js/test_ecma_async.rs

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

async function calc() {
    const left = await 5;
    const right = await 7;
    console.log(left + right);
}
calc();
