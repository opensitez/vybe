// vybe-test: js/async_arrow_functions_lexical_this/test_js_async_arrow_function_returning_object_literal
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

const makeUser = async (name, id) => ({ name, id });
makeUser("Alice", 1).then(user => console.log(`${user.name}:${user.id}`));
