// vybe-test: js/ecma/test_settimeout_nested
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
        setTimeout(() => {
            log.push("first");
            setTimeout(() => {
                log.push("second");
                console.log(log.join(","));
            }, 1);
        }, 1);
