// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_call_apply_bind_cannot_override_this
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

const obj1 = { name: "Obj1" };
const obj2 = { name: "Obj2" };

const fn = async function() {
    // Arrow function inside fn inherits fn's 'this'
    const arrow = async () => this.name;
    return arrow.call(obj2); // call attempt ignored for lexical this
};

fn.call(obj1).then(res => console.log(res));
