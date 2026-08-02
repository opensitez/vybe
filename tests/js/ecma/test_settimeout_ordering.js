// vybe-test: js/ecma/test_settimeout_ordering
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

let order = [];
        setTimeout(() => { order.push("a"); }, 1);
        setTimeout(() => { order.push("b"); }, 2);
        setTimeout(() => { order.push("c"); }, 3);
        setTimeout(() => { console.log(order.join(",")); }, 4);
