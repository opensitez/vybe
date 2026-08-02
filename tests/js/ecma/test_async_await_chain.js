// vybe-test: js/ecma/test_async_await_chain
// origin: languages/js/tests/js/js_ecma_test.rs

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

async function double(x) { return x * 2; }
        async function addTen(x) { return x + 10; }
        
        async function compute() {
            let a = await double(5);
            let b = await addTen(a);
            return b;
        }
        
        console.log(await compute());
