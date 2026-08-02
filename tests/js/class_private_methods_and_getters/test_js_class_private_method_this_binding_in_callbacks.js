// vybe-test: js/class_private_methods_and_getters/test_js_class_private_method_this_binding_in_callbacks
// origin: languages/js/tests/js/test_js_class_private_methods_and_getters.rs

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

class TaskRunner {
    #id = 99;
    #logId() { return `Task_${this.#id}`; }

    execute() {
        const callback = () => this.#logId();
        return callback();
    }
}
__check(__line(new TaskRunner().execute()), "Task_99");
