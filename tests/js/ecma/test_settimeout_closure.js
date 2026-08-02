// vybe-test: js/ecma/test_settimeout_closure
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

let results = [];
        function scheduleOne(val) {
            setTimeout(() => { results.push(val); }, 1);
        }
        for (let i = 0; i < 3; i++) {
            scheduleOne(i);
        }
        setTimeout(() => { console.log(results.join(",")); }, 5);
