// vybe-test: js/do_while_while_for_loop_control_flow/test_js_for_of_loop_return_calls_iterator_return
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
        let i = 0;
        return {
            next() {
                return i < 4 ? { value: ++i, done: false } : { done: true };
            },
            return() {
                returned = true;
                return { done: true };
            }
        };
    }
};
(function consume() {
    for (const n of customIterable) {
        if (n === 2) return n;
    }
}());
console.log(returned);
