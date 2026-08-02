// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_nested_in_class_constructor
// origin: languages/js/tests/js/test_js_async_arrow_functions_lexical_this.rs

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
    constructor(id) {
        this.id = id;
        this.run = async () => `Task_${this.id}`;
    }
}
const t = new Task(99);
t.run().then(res => console.log(res));
