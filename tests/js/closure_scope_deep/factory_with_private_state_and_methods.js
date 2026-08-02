// vybe-test: js/closure_scope_deep/factory_with_private_state_and_methods
// origin: languages/js/tests/js/test_closure_scope_deep.rs

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

function createStack() {
    const items = [];
    return {
        push(v) { items.push(v); return this; },
        pop() { return items.pop(); },
        peek() { return items[items.length - 1]; },
        size() { return items.length; },
        isEmpty() { return items.length === 0; }
    };
}

const s = createStack();
s.push(1).push(2).push(3);
__check(__line(s.size()), "3");
__check(__line(s.peek()), "3");
__check(__line(s.pop()), "3");
__check(__line(s.isEmpty()), "false");
