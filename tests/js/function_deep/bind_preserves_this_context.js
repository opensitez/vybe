// vybe-test: js/function_deep/bind_preserves_this_context
// origin: languages/js/tests/js/test_function_deep.rs

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

const obj = {
    prefix: "Hello",
    greet(name) { return this.prefix + " " + name; }
};
const boundGreet = obj.greet.bind(obj);
const greetAlice = boundGreet.bind(null, "Alice"); // can't override bound this
console.log(greetAlice());
