// vybe-test: js/prototype_oop_patterns/property_delegation
// origin: languages/js/tests/js/test_prototype_oop_patterns.rs

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

function delegate(target, host, methods) {
    for (const m of methods) {
        host[m] = (...args) => target[m](...args);
    }
    return host;
}
class Stack {
    #arr = [];
    push(v) { this.#arr.push(v); return this; }
    pop() { return this.#arr.pop(); }
    peek() { return this.#arr[this.#arr.length - 1]; }
    get size() { return this.#arr.length; }
}
const queue = delegate(new Stack(), {}, ["push", "pop", "peek"]);
queue.push(1); queue.push(2);
console.log(queue.peek());
console.log(queue.pop());
