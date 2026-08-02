// vybe-test: js/do_while_while_for_loop_control_flow/test_js_for_of_loop_iterator_return_called_on_break
// origin: languages/js/tests/js/test_js_do_while_while_for_loop_control_flow.rs

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

let returned = false;
const customIterable = {
    [Symbol.iterator]() {
        return {
            next() { return { value: 1, done: false }; },
            return() { returned = true; return { done: true }; }
        };
    }
};
for (const item of customIterable) {
    break; // Breaking loop closes iterator by calling return()!
}
console.log(returned);
