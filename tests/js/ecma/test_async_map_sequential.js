// vybe-test: js/ecma/test_async_map_sequential
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

async function transform(x) { return x * 10; }
        
        let items = [1, 2, 3];
        let results = [];
        for (let item of items) {
            results.push(await transform(item));
        }
        console.log(results.join(","));
