// vybe-test: js/promise_microtasks/multiple_then_callbacks_ordered
// origin: languages/js/tests/js/test_promise_microtasks.rs

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

const log = [];
const p = Promise.resolve();
p.then(() => log.push("1"));
p.then(() => log.push("2"));
p.then(() => log.push("3"));
p.then(() => console.log(log.join(",")));
