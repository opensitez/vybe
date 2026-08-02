// vybe-test: js/ecma/test_async_in_loop
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

async function process(x) { return x * x; }
        
        let results = [];
        let items = [1, 2, 3, 4, 5];
        for (let item of items) {
            let result = await process(item);
            results.push(result);
        }
        console.log(results.join(","));
