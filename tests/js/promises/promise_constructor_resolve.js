// vybe-test: js/promises/promise_constructor_resolve
// origin: languages/js/tests/js/test_promises.rs

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

let p = new Promise((resolve, reject) => {
    resolve("hello");
});
p.then(v => console.log(v));
