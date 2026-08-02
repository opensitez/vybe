// vybe-test: js/async_error_handling/async_method_in_class
// origin: languages/js/tests/js/test_async_error_handling.rs

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

class DataService {
    async fetch(id) {
        await Promise.resolve();
        return { id, data: "result:" + id };
    }
}
const svc = new DataService();
svc.fetch(42).then(r => console.log(r.data));
