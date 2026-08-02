// vybe-test: js/ecma/test_mixed_sync_async
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

let log = [];
        log.push("1-sync");
        
        async function asyncOp() {
            log.push("2-async-start");
            let result = await Promise.resolve("done");
            log.push("3-async-end");
            return result;
        }
        
        let result = await asyncOp();
        log.push("4-after-await");
        console.log(log.join(" | "));
