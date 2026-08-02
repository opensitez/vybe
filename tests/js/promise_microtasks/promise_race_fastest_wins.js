// vybe-test: js/promise_microtasks/promise_race_fastest_wins
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

const fast = Promise.resolve("fast");
const slow = new Promise(r => setTimeout(() => r("slow"), 100));
Promise.race([fast, slow]).then(v => console.log(v));
