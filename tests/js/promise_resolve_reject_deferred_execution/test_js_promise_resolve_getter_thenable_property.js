// vybe-test: js/promise_resolve_reject_deferred_execution/test_js_promise_resolve_getter_thenable_property
// origin: languages/js/tests/js/test_js_promise_resolve_reject_deferred_execution.rs

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

let getCount = 0;
const obj = {
    get then() {
        getCount++;
        return (resolve) => resolve("GetterThenable");
    }
};
Promise.resolve(obj).then(res => console.log(res + "|GetCount=" + getCount));
