// vybe-test: js/class_private_advanced/private_array_field_public_push_get
// origin: languages/js/tests/js/test_class_private_advanced.rs

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

class Stack {
    #items = [];
    push(item) { this.#items.push(item); }
    pop() { return this.#items.pop(); }
    size() { return this.#items.length; }
    peek() { return this.#items[this.#items.length - 1]; }
}
const s = new Stack();
s.push(10);
s.push(20);
s.push(30);
__check(__line(s.size()), "3");
__check(__line(s.peek()), "30");
__check(__line(s.pop()), "30");
__check(__line(s.size()), "2");
