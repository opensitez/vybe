// vybe-test: js/promise_with_resolvers_utility/test_js_promise_with_resolvers_in_class_field_initializer
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

class Task {
    deferred = Promise.withResolvers();
    finish(data) { this.deferred.resolve(data); }
}
const task = new Task();
task.finish("DoneData");
(async () => {
    console.log(await task.deferred.promise);
})();
