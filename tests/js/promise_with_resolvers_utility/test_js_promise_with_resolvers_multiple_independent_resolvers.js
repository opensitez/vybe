// vybe-test: js/promise_with_resolvers_utility/test_js_promise_with_resolvers_multiple_independent_resolvers
// origin: languages/js/tests/js/test_js_promise_with_resolvers_utility.rs

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

const r1 = Promise.withResolvers();
const r2 = Promise.withResolvers();
r1.resolve(1);
r2.resolve(2);

(async () => {
    const res = await Promise.all([r1.promise, r2.promise]);
    console.log(res.join(","));
})();
